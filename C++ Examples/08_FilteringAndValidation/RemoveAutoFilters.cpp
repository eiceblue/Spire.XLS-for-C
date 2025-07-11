#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveAutoFilters.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"RemoveAutoFilters_out.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		//Remove the auto filters.
		(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->Clear();

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();

}
