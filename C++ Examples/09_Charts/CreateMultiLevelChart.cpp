#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"CreateMultiLevelChart.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Write data to cells
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Main Category");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetText(L"Fruit");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6"))->SetText(L"Vegies");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetText(L"Sub Category");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetText(L"Bananas");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetText(L"Oranges");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetText(L"Pears");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetText(L"Grapes");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B6"))->SetText(L"Carrots");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B7"))->SetText(L"Potatoes");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8"))->SetText(L"Celery");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B9"))->SetText(L"Onions");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetText(L"Value");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetValue(L"52");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetValue(L"65");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetValue(L"50");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5"))->SetValue(L"45");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C6"))->SetValue(L"64");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C7"))->SetValue(L"62");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C8"))->SetValue(L"89");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C9"))->SetValue(L"57");

	////Vertically merge cells from A2 to A5, A6 to A9
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A5"))->XlsRange::Merge();
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6:A9"))->XlsRange::Merge();
	sheet->AutoFitColumn(1);
	sheet->AutoFitColumn(2);

	//Add a clustered bar chart to worksheet
	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::BarClustered);
	chart->SetChartTitle(L"Value");
	dynamic_pointer_cast<ChartPlotArea>(chart->GetPlotArea())->GetFill()->SetFillType(ShapeFillType::NoFill);
	chart->GetLegend()->Delete();
	chart->SetLeftColumn(5);
	chart->SetTopRow(1);
	chart->SetRightColumn(14);

	//Set the data source of series data
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C9")));
	chart->SetSeriesDataFromRange(false);
	//Set the data source of category labels
	intrusive_ptr<ChartSerie> serie = chart->GetSeries()->Get(0);
	serie->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:B9")));
	//Show multi-level category labels
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->SetMultiLevelLable(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
