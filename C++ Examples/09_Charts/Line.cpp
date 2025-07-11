#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"Line.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	sheet->SetName(L"Line Chart");

	//Set chart data
	//Set value of specified cell
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Country");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"Cuba");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"Mexico");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"France");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"German");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"Jun");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(3300);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(2300);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(4500);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(6700);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetValue(L"Jul");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetNumberValue(7500);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetNumberValue(2900);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetNumberValue(2300);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetNumberValue(4200);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D1"))->SetValue(L"Aug");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2"))->SetNumberValue(7400);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D3"))->SetNumberValue(6900);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D4"))->SetNumberValue(7800);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D5"))->SetNumberValue(4200);


	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E1"))->SetValue(L"Sep");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E2"))->SetNumberValue(8000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E3"))->SetNumberValue(7200);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E4"))->SetNumberValue(8500);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E5"))->SetNumberValue(5600);

	//Style
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E1"))->SetRowHeight(15);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E1"))->GetStyle()->SetColor(Spire::Xls::Color::GetDarkGray());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E1"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetWhite());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E1"))->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E1"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D5"))->GetStyle()->SetNumberFormat(L"\"$\"#,##0");

	//Add a chart
	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();
	chart->SetChartType(ExcelChartType::Line);

	//Set region of chart data
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:E5")));

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(11);
	chart->SetBottomRow(29);

	//Set chart title
	chart->SetChartTitle(L"Sales market by country");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Month");
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetIsBold(true);
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetTitleArea()->SetIsBold(true);

	chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
	chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetTextRotationAngle(90);
	chart->GetPrimaryValueAxis()->SetMinValue(1000);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
		cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	}

	dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetVisible(false);

	chart->GetLegend()->SetPosition(LegendPositionType::Top);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
