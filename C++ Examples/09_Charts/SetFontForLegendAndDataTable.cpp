#include "pch.h"
using namespace Spire::Xls;

int main() {
    wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring outputFile = output_path + L"SetFontForLegendAndDataTable.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Create a font with specified size and color
	intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
	font->SetSize(14.0);
	font->SetColor(Spire::Xls::Color::GetRed());

	//Apply the font to chart Legend
	(dynamic_pointer_cast<ChartTextArea>(chart->GetLegend()->GetTextArea()))->SetFont(font);

	//Apply the font to chart DataLabel
	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
		(dynamic_pointer_cast<XlsChartDataLabels>(cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()))->GetTextArea()->SetFont(font);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
