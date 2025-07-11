#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"FormatAxis.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	sheet->SetName(L"FormatAxis");

	//Set chart data
	//Set value of specified cell
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Month");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"Jan");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"Feb");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"Mar");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"Apr");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6"))->SetValue(L"May");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetValue(L"Jun");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8"))->SetValue(L"Jul");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A9"))->SetValue(L"Aug");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"Planned");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(38);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(47);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(39);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(36);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B6"))->SetNumberValue(27);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B7"))->SetNumberValue(25);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8"))->SetNumberValue(36);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B9"))->SetNumberValue(48);

	//Add a chart
	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::ColumnClustered);
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:B9")));
	chart->SetSeriesDataFromRange(false);
	dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->SetVisible(false);
	chart->SetTopRow(10);
	chart->SetBottomRow(28);
	chart->SetLeftColumn(2);
	chart->SetRightColumn(10);
	chart->SetChartTitle(L"Chart with Customized Axis");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);
	intrusive_ptr<ChartSerie> cs1 = chart->GetSeries()->Get(0);
	cs1->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A9")));

	//Format axis
	chart->GetPrimaryValueAxis()->SetMajorUnit(8);
	chart->GetPrimaryValueAxis()->SetMinorUnit(2);
	chart->GetPrimaryValueAxis()->SetMaxValue(50);
	chart->GetPrimaryValueAxis()->SetMinValue(0);
	chart->GetPrimaryValueAxis()->SetIsReverseOrder(false);
	chart->GetPrimaryValueAxis()->SetMajorTickMark(TickMarkType::TickMarkOutside);
	chart->GetPrimaryValueAxis()->SetMinorTickMark(TickMarkType::TickMarkInside);
	chart->GetPrimaryValueAxis()->SetTickLabelPosition(TickLabelPositionType::TickLabelPositionNextToAxis);
	chart->GetPrimaryValueAxis()->SetCrossesAt(0);

	//Set NumberFormat
	chart->GetPrimaryValueAxis()->SetNumberFormat(L"$#,##0");
	chart->GetPrimaryValueAxis()->SetIsSourceLinked(false);

	intrusive_ptr<ChartSerie> serie = chart->GetSeries()->Get(0);
	intrusive_ptr<Spire::Xls::IEnumerator<XlsChartDataPoint>> ie = dynamic_pointer_cast<ChartDataPointsCollection>(serie->GetDataPoints())->GetEnumerator();
	while (ie->MoveNext())
	{
		intrusive_ptr<IChartDataPoint> dataPoint = ie->GetCurrent();
		//Format Series
		dataPoint->GetDataFormat()->GetFill()->SetFillType(ShapeFillType::SolidColor);
		dataPoint->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetLightGreen());

		//Set transparency
		dataPoint->GetDataFormat()->GetFill()->SetTransparency(0.3);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
