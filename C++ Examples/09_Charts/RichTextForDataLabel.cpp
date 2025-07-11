#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartToImage.xlsx";
	wstring outputFile = output_path + L"RichTextForDataLabel.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the first chart inside this worksheet
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Get the first datalabel of the first series 
	intrusive_ptr<IChartDataLabels> datalabel = chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels();

	//Set the text
	datalabel->SetText(L"Rich Text Label");

	//Show the value
	chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels()->SetHasValue(true);

	//Set styles for the text
	//chart.Series[0].DataPoints[0].DataLabels.Font.Color = Color.Red;
	//chart.Series[0].DataPoints[0].DataLabels.Font.IsBold = true;
	chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels()->SetColor(Spire::Xls::Color::GetRed());
	chart->GetSeries()->Get(0)->GetDataPoints()->Get(0)->GetDataLabels()->SetIsBold(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
