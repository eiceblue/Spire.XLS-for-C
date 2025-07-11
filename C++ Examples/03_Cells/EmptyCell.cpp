#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_1.xlsx";
	wstring outputFile = output_path + L"EmptyCell_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set the value as null to remove the original content from the Excel Cell.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C6"))->SetValue(L"");

	//Clear the contents to remove the original content from the Excel Cell.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B6"))->ClearContents();

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D6"))->ClearAll();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
