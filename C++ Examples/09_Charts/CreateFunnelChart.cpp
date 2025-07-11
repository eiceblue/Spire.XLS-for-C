#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"FunnelData.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
	auto sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Add a new chart to the worksheet.
	auto officeChart = sheet->GetCharts()->Add();

	// Set the chart type to Funnel.
	officeChart->SetChartType(ExcelChartType::Funnel);

	// Set the data range for the chart to range A1:B6.
	officeChart->SetDataRange(sheet->GetRange(L"A1:B6"));

	// Set the title of the chart.
	officeChart->SetChartTitle(L"Funnel");

	// Disable the legend in the chart.
	officeChart->SetHasLegend(false);

	// Enable data labels for the default data point of the first series.
	officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);

	// Set the size of the data labels to 8 points.
	officeChart->GetSeries()->Get(0)->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(8);

	// Save the workbook to the specified output file in Excel 2010 format.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);

	// Dispose of the workbook object.
	workbook->Dispose();
}

