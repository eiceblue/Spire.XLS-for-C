#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"ClusteredColumn.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	sheet->SetName(L"ClusteredColumn");

	//Set chart data
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Country");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"Cuba");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"Mexico");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"France");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"German");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"Jun");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(6000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(8000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(9000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(8500);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetValue(L"Aug");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetNumberValue(3000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetNumberValue(2000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetNumberValue(2300);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetNumberValue(4200);

	//Style
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C1"))->SetRowHeight(15);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C1"))->GetStyle()->SetColor(Spire::Common::Color::GetDarkGray());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C1"))->GetStyle()->GetFont()->SetColor(Spire::Common::Color::GetWhite());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C1"))->GetStyle()->SetVerticalAlignment(VerticalAlignType::Center);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C1"))->GetStyle()->SetHorizontalAlignment(HorizontalAlignType::Center);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:C5"))->GetStyle()->SetNumberFormat(L"\"$\"#,##0");

	//Add a chart to the sheet
	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();

	//Set data range of chart 
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5")));
	chart->SetSeriesDataFromRange(false);

	//Set position of the chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(11);
	chart->SetBottomRow(29);

	chart->SetChartType(ExcelChartType::ColumnClustered);

	//Chart title
	chart->SetChartTitle(L"Sales market by country");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	//Chart Axis
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetTitle(L"Country");
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetIsBold(true);
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetTitleArea()->SetIsBold(true);

	chart->GetPrimaryValueAxis()->SetTitle(L"Sales(in Dollars)");
	chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);
	chart->GetPrimaryValueAxis()->SetMinValue(1000);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetIsBold(true);
	chart->GetPrimaryValueAxis()->GetTitleArea()->SetTextRotationAngle(90);

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
		cs->GetFormat()->GetOptions()->SetIsVaryColor(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	}

	//Chart Legend
	chart->GetLegend()->SetPosition(LegendPositionType::Top);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
