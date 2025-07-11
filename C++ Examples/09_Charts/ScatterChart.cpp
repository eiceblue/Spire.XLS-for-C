#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"ScatterChart.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	sheet->SetName(L"Scatter Chart");

	//Set chart data
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

	//Add a chart
	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::ScatterMarkers);

	//Set region of chart data
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B7")));
	chart->SetSeriesDataFromRange(false);

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(11);
	chart->SetRightColumn(10);
	chart->SetBottomRow(28);

	chart->SetChartTitle(L"Scatter Chart");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	chart->GetSeries()->Get(0)->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A7")));
	chart->GetSeries()->Get(0)->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B7")));

	//Add a trend line for the first series
	chart->GetSeries()->Get(0)->GetTrendLines()->Add(TrendLineType::Exponential);

	chart->GetPrimaryValueAxis()->SetTitle(L"Month");
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Planned");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
