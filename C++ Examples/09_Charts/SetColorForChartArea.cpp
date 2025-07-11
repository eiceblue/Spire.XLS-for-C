#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SampeB_4.xlsx";
	wstring outputFile = output_path + L"SetColorForChartArea.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the chart
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Set color for chart area
	dynamic_pointer_cast<ChartArea>(chart->GetChartArea())->GetFill()->SetForeColor(Spire::Xls::Color::GetLightSeaGreen());

	//Set color for plot area
	dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetForeColor(Spire::Xls::Color::GetLightGray());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
