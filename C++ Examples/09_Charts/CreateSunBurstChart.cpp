#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"SunBurstData.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
	auto sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Add a new chart to the worksheet.
	auto officeChart = sheet->GetCharts()->Add();

	// Set the chart type to Sunburst.
	officeChart->SetChartType(ExcelChartType::SunBurst);

	// Set the data range for the chart to range A1:D16.
	officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D16")));

	// Set the top row, bottom row, left column, and right column of the chart's display area.
	officeChart->SetTopRow(1);
	officeChart->SetBottomRow(17);
	officeChart->SetLeftColumn(6);
	officeChart->SetRightColumn(14);

	// Set the title of the chart.
	officeChart->SetChartTitle(L"Sales by quarter");

	// Set the size of the data labels for the default data point of the first series to 8 points.
	officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(8);

	// Disable the legend in the chart.
	officeChart->SetHasLegend(false);

	// Save the workbook to the specified output file in Excel 2010 format.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);

	// Dispose of the workbook object.
	workbook->Dispose();
}

