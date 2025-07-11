#include "pch.h"
using namespace Spire::Xls;

int main() {
    wstring output_path = OUTPUTPATH;
    wstring outputFile = output_path + L"AddErrorBars_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Create a empty sheet
	workbook->CreateEmptySheets(1);

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add data
	sheet->SetName(L"Demo");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Month");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"Jan.");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"Feb.");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"Mar.");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"Apr.");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6"))->SetValue(L"May.");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetValue(L"Jun.");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"Planned");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(3.3);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(2.5);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(2.0);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(3.7);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B6"))->SetNumberValue(4.5);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B7"))->SetNumberValue(4.0);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetValue(L"Actual");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetNumberValue(3.8);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetNumberValue(3.2);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetNumberValue(1.7);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetNumberValue(3.5);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C6"))->SetNumberValue(4.5);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C7"))->SetNumberValue(4.3);

	//Add a line chart and then add percentage error bar to the chart
	intrusive_ptr<Spire::Xls::Chart> chart = sheet->GetCharts()->Add(ExcelChartType::Line);
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:B7")));
	chart->SetSeriesDataFromRange(false);
	//Set chart position
	chart->SetTopRow(8);
	chart->SetBottomRow(25);
	chart->SetLeftColumn(2);
	chart->SetRightColumn(9);
	chart->SetChartTitle(L"Error Bar 10% Plus");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);
	intrusive_ptr<ChartSerie> cs1 = chart->GetSeries()->Get(0);
	cs1->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A7")));
	cs1->ErrorBar(true, ErrorBarIncludeType::Plus, ErrorBarType::Percentage, 10);

	// Add a column chart with standard error bars as comparison
	intrusive_ptr<Spire::Xls::Chart> chart2 = dynamic_pointer_cast<Spire::Xls::Chart>(sheet->GetCharts()->Add(ExcelChartType::ColumnClustered));
	chart2->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:C7")));
	chart2->SetSeriesDataFromRange(false);
	//Set chart position
	chart2->SetTopRow(8);
	chart2->SetBottomRow(25);
	chart2->SetLeftColumn(10);
	chart2->SetRightColumn(17);
	chart2->SetChartTitle(L"Standard Error Bar");
	chart2->GetChartTitleArea()->SetIsBold(true);
	chart2->GetChartTitleArea()->SetSize(12);
	intrusive_ptr<ChartSerie> cs2 = chart2->GetSeries()->Get(0);
	cs2->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A7")));
	cs2->ErrorBar(true, ErrorBarIncludeType::Minus, ErrorBarType::StandardError, 0.3);
	intrusive_ptr<ChartSerie> cs3 = chart2->GetSeries()->Get(1);
	cs3->ErrorBar(true, ErrorBarIncludeType::Both, ErrorBarType::StandardError, 0.5);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
