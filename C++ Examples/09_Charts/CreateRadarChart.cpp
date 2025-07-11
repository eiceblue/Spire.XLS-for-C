#include "pch.h"
using namespace Spire::Xls;

int main() {	
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"CreateRadarChart.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Initailize worksheet
	workbook->CreateEmptySheets(1);
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	sheet->SetName(L"Chart data");
	sheet->SetGridLinesVisible(false);

	//Writes chart data
	//Product
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Product");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"Bikes");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"Cars");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"Trucks");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"Buses");

	//Paris
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"Paris");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(4000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(23000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(4000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(30000);

	//New York
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetValue(L"New York");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetNumberValue(30000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetNumberValue(7600);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetNumberValue(18000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetNumberValue(8000);

	//Style
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C1"))->GetStyle()->GetFont()->SetIsBold(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:C2"))->GetStyle()->SetKnownColor(ExcelColors::LightYellow);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3:C3"))->GetStyle()->SetKnownColor(ExcelColors::LightGreen1);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4:C4"))->GetStyle()->SetKnownColor(ExcelColors::LightOrange);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:C5"))->GetStyle()->SetKnownColor(ExcelColors::LightTurquoise);

	//Border
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"))->GetStyle()->GetBorders()->Get(BordersLineType::EdgeTop)->SetColor(Spire::Common::Color::FromArgb(0, 0, 128));
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"))->GetStyle()->GetBorders()->Get(BordersLineType::EdgeTop)->SetLineStyle(LineStyleType::Thin);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"))->GetStyle()->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Common::Color::FromArgb(0, 0, 128));
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"))->GetStyle()->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Thin);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"))->GetStyle()->GetBorders()->Get(BordersLineType::EdgeLeft)->SetColor(Spire::Common::Color::FromArgb(0, 0, 128));
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"))->GetStyle()->GetBorders()->Get(BordersLineType::EdgeLeft)->SetLineStyle(LineStyleType::Thin);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"))->GetStyle()->GetBorders()->Get(BordersLineType::EdgeRight)->SetColor(Spire::Common::Color::FromArgb(0, 0, 128));
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5"))->GetStyle()->GetBorders()->Get(BordersLineType::EdgeRight)->SetLineStyle(LineStyleType::Thin);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:C5"))->GetStyle()->SetNumberFormat(L"\"$\"#,##0");
	//Add a new  chart worsheet to workbook
	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();

	//Set position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(6);
	chart->SetRightColumn(11);
	chart->SetBottomRow(29);

	//Set region of chart data
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C5")));
	chart->SetSeriesDataFromRange(false);

	chart->SetChartType(ExcelChartType::Radar);

	//Chart title
	chart->SetChartTitle(L"Sale market by region");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetVisible(false);

	chart->GetLegend()->SetPosition(LegendPositionType::Corner);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
