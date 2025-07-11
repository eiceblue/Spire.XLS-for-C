#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"PivotTableExample.xlsx";
	wstring outputFile = output + L"ClearPivotFields.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	
	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"PivotTable"));

	intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));

	//Clear all the data fields
	pt->GetDataFields()->Clear();

	pt->CalculateData();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}