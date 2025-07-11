#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"PivotTableExample.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create an intrusive pointer to a Workbook object
	intrusive_ptr<Workbook>workbook = new Workbook();

	// Load the workbook from the input file
	workbook->LoadFromFile(inputFile.c_str());

	// Get the second worksheet from the workbook and retrieve the first pivot table
	intrusive_ptr<XlsPivotTable>pivotTable = dynamic_pointer_cast<XlsPivotTable>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetPivotTables()->Get(0));
	
	// Set the option to show the data field in the row for the pivot table
	pivotTable->SetShowDataFieldInRow(true);

	// Calculate the data for the pivot table
	pivotTable->CalculateData();

	// Save the workbook to the output file in Excel 2016 format
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2016);

	// Dispose of the workbook object
	workbook->Dispose();
}

