#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"DataCallout.xlsx";
	wstring outputFile = output_path + L"DataCallout_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the first chart
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasValue(true);
		(dynamic_pointer_cast<XlsChartDataLabels>(cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()))->SetHasWedgeCallout(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasCategoryName(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasSeriesName(true);
		cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()->SetHasLegendKey(true);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
