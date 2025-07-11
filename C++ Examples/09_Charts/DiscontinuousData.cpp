#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"DiscontinuousData.xlsx";
	wstring outputFile = output_path + L"DiscontinuousData_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add a chart
	intrusive_ptr<Spire::Xls::Chart> chart = sheet->GetCharts()->Add(ExcelChartType::ColumnClustered);
	chart->SetSeriesDataFromRange(false);

	//Set the position of chart
	chart->SetLeftColumn(1);
	chart->SetTopRow(10);
	chart->SetRightColumn(10);
	chart->SetBottomRow(24);

	//Add a series
	intrusive_ptr<ChartSerie> cs1 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Add());

	//Set the name of the cs1
	cs1->SetName(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetValue());

	//Set discontinuous values for cs1
	cs1->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A3"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:A6"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:A9")))));
	cs1->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:B3"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5:B6"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8:B9")))));

	//Set the chart type
	cs1->SetSerieType(ExcelChartType::ColumnClustered);

	//Add a series
	intrusive_ptr<ChartSerie> cs2 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Add());
	cs2->SetName(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->GetValue());
	cs2->SetCategoryLabels(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:A3"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5:A6"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:A9")))));
	cs2->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C3"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C5:C6"))->AddCombinedRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C8:C9")))));
	cs2->SetSerieType(ExcelChartType::ColumnClustered);

	chart->SetChartTitle(L"Chart");
	chart->GetChartTitleArea()->GetFont()->SetSize(20);
	chart->GetChartTitleArea()->SetColor(Spire::Xls::Color::GetBlack());

	chart->GetPrimaryValueAxis()->SetHasMajorGridLines(false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
