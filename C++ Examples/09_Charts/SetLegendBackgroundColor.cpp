#include "pch.h"
using namespace Spire::Xls;

int main() {
                wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ChartSample1.xlsx";
	wstring outputFile = output_path + L"SetLegendBackgroundColor.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	intrusive_ptr<XlsChartFrameFormat> x = dynamic_pointer_cast<XlsChartFrameFormat>(dynamic_pointer_cast<ChartLegend>(chart->GetLegend())->GetFrameFormat());
	x->GetFill()->SetFillType(ShapeFillType::SolidColor);
	x->SetForeGroundColor(Spire::Xls::Color::GetSkyBlue());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
