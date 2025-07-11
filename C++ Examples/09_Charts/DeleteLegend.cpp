#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring outputFile = output_path + L"DeleteLegend.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the chart
	intrusive_ptr<Spire::Xls::Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	////Delete legend from the chart
	//chart.Legend.Delete();

	//Delete the first and the second legend entries from the chart
	chart->GetLegend()->GetLegendEntries()->Get(0)->Delete();
	chart->GetLegend()->GetLegendEntries()->Get(1)->Delete();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
