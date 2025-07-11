#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"PivotTableExample.xlsx";
	wstring outputFile = output + L"UpdateDataSource.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> data = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"Data"));
	intrusive_ptr<CellRange> a2 = dynamic_pointer_cast<CellRange>(data->Get(L"A2"));

	a2->SetText(L"NewValue");
	data->Get(L"D2")->SetNumberValue(28000);

	//Get the sheet in which the pivot table is located
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(L"PivotTable"));

	intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));
	//Refresh and calculate
	pt->GetCache()->SetIsRefreshOnLoad(true);
	pt->CalculateData();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}