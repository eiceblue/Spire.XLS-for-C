#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring outputFile = output_path + L"ApplySoftEdgesEffect.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the chart
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Specify the size of the soft edge. Value can be set from 0 to 100
	dynamic_pointer_cast<ChartArea>(chart->GetChartArea())->GetShadow()->SetSoftEdge(25);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
