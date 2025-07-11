#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"Pie.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	sheet->SetName(L"Pie Chart");

	//Add a chart
	intrusive_ptr<Chart> chart = nullptr;
	chart = sheet->GetCharts()->Add(ExcelChartType::Pie);

	//Set chart data
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Year");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"2002");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"2003");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"2004");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"2005");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"Sales");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(4000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(6000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(7000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(8500);

	//Style
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B1"))->SetRowHeight(15);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B1"))->GetStyle()->SetColor(Spire::Xls::Color::GetDarkGray());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B1"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetWhite());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B1"))->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B1"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:C5"))->GetStyle()->SetNumberFormat(L"\"$\"#,##0");

	//Set region of chart data
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B5")));
	chart->SetSeriesDataFromRange(false);

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(9);
	chart->SetBottomRow(25);

	//Chart title
	chart->SetChartTitle(L"Sales by year");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(0);
	cs->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A5")));
	cs->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B5")));
	cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);

	dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetVisible(false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
