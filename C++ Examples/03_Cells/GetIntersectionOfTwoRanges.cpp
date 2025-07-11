#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_1.xlsx";
	wstring outputFile = output_path + L"GetIntersectionOfTwoRanges_result.txt";

	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the two ranges.
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:D7"))->Intersect(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:E8")));

	std::wstring* content = new std::wstring();
	content->append(L"The intersection of the two ranges \"A2:D7\" and \"B2:E8\" is:");

	//Get the intersection of the two ranges.
	//for (auto r : range->GetCells())
	//{
	for (int i = 0; i < range->GetCells()->GetCount(); i++)
	{
		intrusive_ptr<CellRange> cr = range->GetCells()->GetItem(i);
		content->append(cr->GetValue());
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << content << endl;
	ofs.close();
	workbook->Dispose();
}

