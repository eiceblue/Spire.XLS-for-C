#include "pch.h"
using namespace Spire::Xls;

void main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Test.xlam";
	wstring outputFile = output_path + L"RegisterAddInFunction.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Register AddIn function
	workbook->GetAddInFunctions()->Add(inputFile.c_str(), L"TEST_UDF");
	workbook->GetAddInFunctions()->Add(inputFile.c_str(), L"TEST_UDF1");
	//Get the first sheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Call AddIn function
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetFormula(L"=TEST_UDF()");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetFormula(L"=TEST_UDF1()");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}