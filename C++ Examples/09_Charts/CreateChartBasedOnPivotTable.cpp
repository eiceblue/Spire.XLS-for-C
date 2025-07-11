#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"PivotTable.xlsx";
	wstring outputFile = output_path + L"CreateChartBasedOnPivotTable.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	intrusive_ptr<XlsPivotTable> pt = dynamic_pointer_cast<XlsPivotTable>(sheet->GetPivotTables()->Get(0));

	dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetCharts()->Add(ExcelChartType::BarClustered, pt);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

}
