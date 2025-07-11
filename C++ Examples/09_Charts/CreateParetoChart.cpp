#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"ParetoChartData.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Retrieve the first worksheet from the workbook using dynamic_pointer_cast.
	auto sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Add a new chart to the worksheet.
	auto officeChart = sheet->GetCharts()->Add();

	// Set the chart type to Pareto.
	officeChart->SetChartType(ExcelChartType::Pareto);

	// Set the data range for the chart to range A2:B8.
	officeChart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:B8")));

	// Set the top row, bottom row, left column, and right column of the chart's display area.
	officeChart->SetTopRow(1);
	officeChart->SetBottomRow(19);
	officeChart->SetLeftColumn(4);
	officeChart->SetRightColumn(12);

	// Get the primary category axis (X-axis) of the chart and cast it to ChartCategoryAxis.
	auto axis = dynamic_pointer_cast<ChartCategoryAxis>(officeChart->GetPrimaryCategoryAxis());

	// Enable binning by category on the primary category axis.
	axis->SetIsBinningByCategory(true);

	// Set the value to be used for overflow bins on the primary category axis.
	axis->SetOverflowBinValue(5);

	// Set the value to be used for underflow bins on the primary category axis.
	axis->SetUnderflowBinValue(1);

	// Set the line color of the Pareto line in the chart.
	officeChart->GetSeries()->Get(0)->GetParetoLineFormat()->GetLineProperties()->SetColor(Color::GetBlue());

	// Set the gap width between bars in the chart's series.
	officeChart->GetSeries()->Get(0)->GetDataFormat()->GetOptions()->SetGapWidth(6);

	// Set the title of the chart.
	officeChart->SetChartTitle(L"Expenses");

	// Disable the legend in the chart.
	officeChart->SetHasLegend(false);

	// Save the workbook to the specified output file in Excel 2010 format.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);

	// Dispose of the workbook object.
	workbook->Dispose();
}

