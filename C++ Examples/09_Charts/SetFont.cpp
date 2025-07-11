#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SetFont.xlsx";
	wstring outputFile = output_path + L"SetFont.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the first sheet
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Create a font
	intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
	font->SetSize(15.0);
	font->SetColor(Spire::Xls::Color::GetLightSeaGreen());

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		intrusive_ptr<ChartSerie> cs = chart->GetSeries()->Get(i);
		//Set font
		(dynamic_pointer_cast<XlsChartDataLabels>(cs->GetDataPoints()->GetDefaultDataPoint()->GetDataLabels()))->GetTextArea()->SetFont(font);
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
