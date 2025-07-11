#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring  inputFile = input_path + L"ChartSample1.xlsx";
	wstring  outputFile = output_path + L"SetFontForTitleAndAxis.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Format the font for the chart title
	chart->GetChartTitleArea()->SetColor(Spire::Xls::Color::GetBlue());
	chart->GetChartTitleArea()->SetSize(20.0);

	//Format the font for the chart Axis
	chart->GetPrimaryValueAxis()->GetFont()->SetColor(Spire::Xls::Color::GetGold());
	chart->GetPrimaryValueAxis()->GetFont()->SetSize(10.0);
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetColor(Spire::Xls::Color::GetRed());
	dynamic_pointer_cast<ChartCategoryAxis>(chart->GetPrimaryCategoryAxis())->GetFont()->SetSize(20.0);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
