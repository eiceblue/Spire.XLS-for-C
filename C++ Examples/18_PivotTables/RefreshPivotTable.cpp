#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;

	wstring inputFile = fn + L"Template_Xls_7.xlsx";
	wstring outputFile = output + L"RefreshPivotTable.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

	//Update the data source of PivotTable.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2"))->SetValue(L"999");

	//Get the PivotTable that was built on the data source.
	intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetPivotTables()->Get(0));

	//Refresh the data of PivotTable.
	pt->GetCache()->SetIsRefreshOnLoad(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}