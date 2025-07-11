#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"GaugeChart.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	sheet->SetName(L"Gauge Chart");

	//Set chart data
	//Set value of specified cell
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Value");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"30");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"60");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"90");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"180");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetValue(L"value");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetValue(L"pointer");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetValue(L"End");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2"))->SetValue(L"10");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D3"))->SetValue(L"1");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D4"))->SetValue(L"189");

	//Add a Doughnut chart
	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::Doughnut);
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:A5")));
	chart->SetSeriesDataFromRange(false);
	chart->SetHasLegend(true);

	//Set the position of chart
	chart->SetLeftColumn(2);
	chart->SetTopRow(7);
	chart->SetRightColumn(9);
	chart->SetBottomRow(25);

	//Get the series 1
	intrusive_ptr<ChartSerie> cs1 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Get(L"Value"));
	cs1->GetFormat()->GetOptions()->SetDoughnutHoleSize(60);
	cs1->GetDataFormat()->GetOptions()->SetFirstSliceAngle(270);

	//Set the fill color
	cs1->GetDataPoints()->Get(0)->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetYellow());
	cs1->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetPaleVioletRed());
	cs1->GetDataPoints()->Get(2)->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetDarkViolet());
	cs1->GetDataPoints()->Get(3)->GetDataFormat()->GetFill()->SetVisible(false);

	//Add a series with pie chart
	intrusive_ptr<ChartSerie> cs2 = static_cast<intrusive_ptr<ChartSerie>>(chart->GetSeries()->Add(L"Pointer", ExcelChartType::Pie));

	//Set the value
	cs2->SetValues(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2:D4")));
	cs2->SetUsePrimaryAxis(false);
	cs2->GetDataPoints()->Get(0)->GetDataLabels()->SetHasValue(true);
	cs2->GetDataFormat()->GetOptions()->SetFirstSliceAngle(270);
	cs2->GetDataPoints()->Get(0)->GetDataFormat()->GetFill()->SetVisible(false);
	cs2->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetFillType(ShapeFillType::SolidColor);
	cs2->GetDataPoints()->Get(1)->GetDataFormat()->GetFill()->SetForeColor(Spire::Xls::Color::GetBlack());
	cs2->GetDataPoints()->Get(2)->GetDataFormat()->GetFill()->SetVisible(false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
