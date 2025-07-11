#include "pch.h"
using namespace Spire::Xls;

int main() { 
	wstring output_path = OUTPUTPATH;

	wstring outputFile = output_path + L"CreateDoughnutChart.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Insert data
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Country");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->GetStyle()->GetFont()->SetIsBold(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"Cuba");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"Mexico");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"France");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"German");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"Sales");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetIsBold(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(6000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(8000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(9000);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"))->SetNumberValue(8500);

	//Add a new chart, set chart type as doughnut
	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add();
	chart->SetChartType(ExcelChartType::Doughnut);
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:B5")));
	chart->SetSeriesDataFromRange(false);

	//Set position of chart
	chart->SetLeftColumn(4);
	chart->SetTopRow(2);
	chart->SetRightColumn(12);
	chart->SetBottomRow(22);

	//Chart title
	chart->SetChartTitle(L"Market share by country");
	chart->GetChartTitleArea()->SetIsBold(true);
	chart->GetChartTitleArea()->SetSize(12);

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasPercentage(true);
	}

	chart->GetLegend()->SetPosition(LegendPositionType::Top);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
