#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Template_Xls_1.xlsx";
	wstring outputFile = outputFolder + L"UnmergeCells_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Unmerge the cells.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"F2"))->UnMerge();

	//Unmerge the cells.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"F7"))->UnMerge();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
