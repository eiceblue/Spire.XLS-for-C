#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"BoxAndWhiskerChartData.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
	auto sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Add a new chart to the worksheet.
	auto officeChart = sheet->GetCharts()->Add();

	// Set the title of the chart.
	officeChart->SetChartTitle(L"Yearly Vehicle Sales");

	// Set the chart type to Box and Whisker.
	officeChart->SetChartType(ExcelChartType::BoxAndWhisker);

	// Set the data range for the chart to range A1:E17.
	officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E17")));

	// Get the first series from the chart's series collection.
	auto seriesA = officeChart->GetSeries()->Get(0);
	seriesA->GetDataFormat()->SetShowInnerPoints(false);
	seriesA->GetDataFormat()->SetShowOutlierPoints(true);
	seriesA->GetDataFormat()->SetShowMeanMarkers(true);
	seriesA->GetDataFormat()->SetShowMeanLine(false);
	seriesA->GetDataFormat()->SetQuartileCalculationType(ExcelQuartileCalculation::ExclusiveMedian);

	// Get the second series from the chart's series collection.
	auto seriesB = officeChart->GetSeries()->Get(1);
	seriesB->GetDataFormat()->SetShowInnerPoints(false);
	seriesB->GetDataFormat()->SetShowOutlierPoints(true);
	seriesB->GetDataFormat()->SetShowMeanMarkers(true);
	seriesB->GetDataFormat()->SetShowMeanLine(false);
	seriesB->GetDataFormat()->SetQuartileCalculationType(ExcelQuartileCalculation::InclusiveMedian);

	// Get the third series from the chart's series collection.
	auto seriesC = officeChart->GetSeries()->Get(2);
	seriesC->GetDataFormat()->SetShowInnerPoints(false);
	seriesC->GetDataFormat()->SetShowOutlierPoints(true);
	seriesC->GetDataFormat()->SetShowMeanMarkers(true);
	seriesC->GetDataFormat()->SetShowMeanLine(false);
	seriesC->GetDataFormat()->SetQuartileCalculationType(ExcelQuartileCalculation::ExclusiveMedian);

	// Save the workbook to the specified output file in Excel 2010 format.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);

	// Dispose of the workbook object.
	workbook->Dispose();
}

