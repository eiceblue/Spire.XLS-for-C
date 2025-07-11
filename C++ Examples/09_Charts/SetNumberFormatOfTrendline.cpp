#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring  input_path = DATAPATH;
	wstring  output_path = OUTPUTPATH;
	wstring  inputFile = input_path + L"ChartSample4.xlsx";
	wstring  outputFile = output_path + L"SetNumberFormatOfTrendline.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the chart from the first worksheet
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetCharts()->Get(0));

	//Get the trendline of the chart and then extract the equation of the trendline
	intrusive_ptr<IChartTrendLine> trendLine = chart->GetSeries()->Get(1)->GetTrendLines()->GetItem(0);

	//Set the number format of trendLine to "#,##0.00"
	trendLine->GetDataLabel()->SetNumberFormat(L"#,##0.00");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
