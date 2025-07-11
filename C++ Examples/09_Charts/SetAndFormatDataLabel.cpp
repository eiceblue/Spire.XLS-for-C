#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"SetAndFormatDataLabel.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->CreateEmptySheets(1);

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	sheet->SetName(L"Demo");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Month");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"Jan");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"Feb");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"Mar");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"Apr");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6"))->SetValue(L"May");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetValue(L"Jun");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"Peter");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(25);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(18);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(8);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(13);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B6"))->SetNumberValue(22);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B7"))->SetNumberValue(28);

	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::LineMarkers);
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1:B7")));
	dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->SetVisible(false);
	chart->SetSeriesDataFromRange(false);
	chart->SetTopRow(5);
	chart->SetBottomRow(26);
	chart->SetLeftColumn(2);
	chart->SetRightColumn(11);
	chart->SetChartTitle(L"Data Labels Demo");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);
	intrusive_ptr<ChartSerie> cs1 = chart->GetSeries()->Get(0);
	cs1->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A7")));

	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasLegendKey(false);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasPercentage(false);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasSeriesName(true);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasCategoryName(true);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetDelimiter(L". ");

	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetSize(9);
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetColor(Spire::Xls::Color::GetRed());
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetFontName(L"Calibri");
	cs1->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetPosition(DataLabelPositionType::Center);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
