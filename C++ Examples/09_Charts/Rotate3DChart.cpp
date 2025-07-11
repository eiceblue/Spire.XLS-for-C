#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartSample3.xlsx";
	wstring outputFile = output_path + L"Rotate3DChart.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//X rotation:
	chart->SetRotation(30);
	//Y rotation:
	chart->SetElevation(20);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
