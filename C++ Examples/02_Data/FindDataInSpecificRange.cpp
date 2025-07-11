#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"FindCellsSample.xlsx";
	wstring outputFile = output_path + L"FindDataInSpecificRange_result.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Specify a range
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(1, 1, 12, 8));

	//Create a string
	std::wstring* builder = new std::wstring();

	//Find string from this range
	auto textRanges = range->FindAllString(L"E-iceblue", false, false);

	//Append the address of found cells in builder
	if (textRanges->GetCount() != 0)
	{
		//for (auto r : textRanges)
		//{
		for (int i = 0; i < textRanges->GetCount(); i++)
		{
			intrusive_ptr<CellRange> cr = textRanges->GetItem(i);
			wstring address = cr->GetRangeAddress();
			builder->append(L"The address of found text cell is: " + address);
		}
	}
	else
	{
		builder->append(L"No cell contain the text.");
	}


	//Find number from this range
	auto ranges = range->FindAllNumber(100, true);

	//Append the address of found cells in builder
	if (ranges->GetCount() != 0)
	{
		//for (auto r : numberRanges)
		//{
		for (int i = 0; i < ranges->GetCount(); i++)
		{
			intrusive_ptr<CellRange> r = ranges->GetItem(i);
			wstring address = r->GetRangeAddress();
			builder->append(L"The address of found number cell is: " + address);
		}
	}
	else
	{
		builder->append(L"No cell contain the number.");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << builder << endl;
	ofs.close();
	workbook->Dispose();

}

