#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_7.xlsx";
	wstring outputFile = output + L"ExpandOrCollapseRows.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the data in Pivot Table.
	intrusive_ptr<XlsPivotTable> pivotTable = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));

	//Calculate Data.
	pivotTable->CalculateData();

	//Collapse the rows.
	(dynamic_pointer_cast<XlsPivotField>(pivotTable->GetPivotFields()->Get(L"Vendor No")))->HideItemDetail(L"3501", true);

	//Expand the rows.
	(dynamic_pointer_cast<XlsPivotField>(pivotTable->GetPivotFields()->Get(L"Vendor No")))->HideItemDetail(L"3502", false);


	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}