#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"FormatDataField.xlsx";
	wstring outputFile = output + L"FormatDataField.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Access the PivotTable
	intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));
	// Access the data field.
	intrusive_ptr<PivotDataField> pivotDataField = pt->GetDataFields()->Get(0);
	// Set data display format
	pivotDataField->SetShowDataAs(PivotFieldFormatType::PercentageOfColumn);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}