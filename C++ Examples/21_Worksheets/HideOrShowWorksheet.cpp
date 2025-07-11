#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"WorksheetSample3.xlsx";
	std::wstring outputFile = output_path + L"HideOrShowWorksheet.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Hide the sheet named "Sheet1"
	dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"Sheet1"))->SetVisibility(WorksheetVisibility::Hidden);
	//Show the second sheet
	dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->SetVisibility(WorksheetVisibility::Visible);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}