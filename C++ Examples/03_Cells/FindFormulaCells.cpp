#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"FindCellsSample.xlsx";
	wstring outputFile = output_path + L"FindFormulaCells_result.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Find the cells that contain formula L"=SUM(A11,A12)"
	intrusive_ptr<Spire::Xls::IList<CellRange>> ranges = sheet->FindAll(L"=SUM(A11,A12)", FindType::Formula, ExcelFindOptions::None);

	//Create a string* builder = new std::wstring()
	std::wstring* builder = new std::wstring();

	//Append the address of found cells to builder
	if (ranges->GetCount() != 0)
	{
		//for (auto range : ranges)
		//{
		for (int i = 0; i < ranges->GetCount(); i++)
		{
			intrusive_ptr<CellRange> cr = ranges->GetItem(i);
			wstring address = cr->GetRangeAddress();
			builder->append(L"The address of found cell is: " + address);
		}
	}
	else
	{
		builder->append(L"No cell contain the formula");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << builder << endl;
	ofs.close();
	workbook->Dispose();

}

