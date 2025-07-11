#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ShowLeaderLine.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set value of specified range
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"1");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"2");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"3");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetValue(L"4");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetValue(L"5");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetValue(L"6");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetValue(L"7");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetValue(L"8");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetValue(L"9");

	intrusive_ptr<Chart> chart = sheet->GetCharts()->Add(ExcelChartType::BarStacked);
	chart->SetDataRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:C3")));
	chart->SetTopRow(4);
	chart->SetLeftColumn(2);
	chart->SetWidth(450);
	chart->SetHeight(300);

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetShowLeaderLines(true);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
	
}
