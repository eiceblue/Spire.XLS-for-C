#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"ChangeDataSource.xlsx";
	wstring outputFile = output + L"ChangeDataSource.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	intrusive_ptr<CellRange> Range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C15"));
	intrusive_ptr<XlsPivotTable> table = dynamic_pointer_cast<XlsPivotTable>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetPivotTables()->Get(0));

	//Change data source
	table->ChangeDataSource(Range);
	table->GetCache()->SetIsRefreshOnLoad(false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

