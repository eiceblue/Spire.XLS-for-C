#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"CreateTable.xlsx";
	wstring outputFile = output_path + L"CopyCellsRange_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Specify a destination range 
	intrusive_ptr<CellRange> cells = dynamic_pointer_cast<CellRange>(sheet1->GetRange(L"G1:H19"));

	//Copy the selected range to destination range 
	dynamic_pointer_cast<CellRange>(sheet1->GetRange(L"B1:C19"))->Copy(cells);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

