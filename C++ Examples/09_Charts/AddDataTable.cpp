#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AddDataTable.xlsx";
    wstring output_path = OUTPUTPATH;
    wstring outputFile = output_path + L"AddDataTable_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the first chart
	intrusive_ptr<Spire::Xls::Chart> chart = dynamic_pointer_cast<Spire::Xls::Chart>(dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0)));
	chart->SetHasDataTable(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
	workbook->Dispose();
}
