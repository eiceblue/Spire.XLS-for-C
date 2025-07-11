#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"HistogramChartData.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
	auto sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Add a new chart to the worksheet and cast it to Chart type.
	auto officeChart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Add());

	// Set the chart type to Histogram.
	officeChart->SetChartType(ExcelChartType::Histogram);

	// Set the data range for the chart to column A (A1:A15).
	officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:A15")));

	// Set the top row, bottom row, left column, and right column of the chart's display area.
	officeChart->SetTopRow(1);
	officeChart->SetBottomRow(19);
	officeChart->SetLeftColumn(4);
	officeChart->SetRightColumn(12);

	// Set the bin width for the chart's primary category axis (X-axis).
	dynamic_pointer_cast<ChartCategoryAxis>(officeChart->GetPrimaryCategoryAxis())->SetBinWidth(8);

	// Set the gap width between bars in the chart's series.
	officeChart->GetSeries()->Get(0)->GetDataFormat()->GetOptions()->SetGapWidth(6);

	// Set the title of the chart.
	officeChart->SetChartTitle(L"Height Data");

	// Set the title of the primary value axis (Y-axis).
	officeChart->GetPrimaryValueAxis()->SetTitle(L"Number of students");

	// Set the title of the primary category axis (X-axis).
	officeChart->GetPrimaryCategoryAxis()->SetTitle(L"Height");

	// Disable the legend in the chart.
	officeChart->SetHasLegend(false);

	// Save the workbook to the specified output file in Excel 2010 format.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);

	// Dispose of the workbook object.
	workbook->Dispose();
}

