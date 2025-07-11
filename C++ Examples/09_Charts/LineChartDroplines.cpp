#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"LineChartDroplines.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create an intrusive pointer to a Workbook object
	intrusive_ptr<Workbook>workbook = new Workbook();

	// Load the workbook from the input file
	workbook->LoadFromFile(inputFile.c_str());

	// Get the first worksheet from the workbook
	intrusive_ptr<Worksheet>worksheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Get the first chart from the worksheet
	intrusive_ptr<Chart>chart = dynamic_pointer_cast<Chart>(worksheet->GetCharts()->Get(0));

	// Enable droplines for the first series of the chart
	chart->GetSeries()->Get(0)->SetHasDroplines(true);

	// Save the workbook to the output file in Excel 2013 format
	workbook->SaveToFile(outputFile.c_str(), FileFormat::Version2013);

	// Dispose of the workbook object
	workbook->Dispose();
}

