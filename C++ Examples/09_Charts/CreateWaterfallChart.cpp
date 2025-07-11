#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"WaterfallChartData.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create an intrusive pointer to a Workbook object
	intrusive_ptr<Workbook>workbook = new Workbook();

	// Load the workbook from the input file
	workbook->LoadFromFile(inputFile.c_str());

	// Get the first worksheet from the workbook
	auto sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Add a chart to the worksheet
	auto officeChart = sheet->GetCharts()->Add();

	// Set the chart type to Waterfall
	officeChart->SetChartType(ExcelChartType::WaterFall);

	// Set the data range for the chart
	officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:B8")));

	// Set the top row, bottom row, left column, and right column for the chart
	officeChart->SetTopRow(1);
	officeChart->SetBottomRow(19);
	officeChart->SetLeftColumn(4);
	officeChart->SetRightColumn(12);

	// Set certain data points in the chart as total
	dynamic_pointer_cast<XlsChartDataPoint>(officeChart->GetSeries()->Get(0)->GetDataPoints()->Get(3))->SetSetAsTotal(true);
	dynamic_pointer_cast<XlsChartDataPoint>(officeChart->GetSeries()->Get(0)->GetDataPoints()->Get(6))->SetSetAsTotal(true);
	
	// Show connector lines in the chart
	dynamic_pointer_cast<ChartSerieDataFormat>(officeChart->GetSeries()->Get(0)->GetFormat())->SetShowConnectorLines(true);
	
	// Set the chart title
	officeChart->SetChartTitle(L"WaterFall Chart");

	// Enable data labels for the default data point of the chart series
	officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(8);

	// Set the position of the legend to the right
	officeChart->GetLegend()->SetPosition(LegendPositionType::Right);

	// Save the workbook to the output file in Excel 2010 format
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);

	// Dispose of the workbook object
	workbook->Dispose();
}

