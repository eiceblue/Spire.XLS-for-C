#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AutofilterBlank.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"AutofilterNonBlank_out.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		//Match the non blank data
		(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->MatchNonBlanks(0);

		//Filter
		(dynamic_pointer_cast<AutoFiltersCollection>(sheet->GetAutoFilters()))->Filter();

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();
}
