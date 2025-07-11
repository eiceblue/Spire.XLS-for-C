#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CSVToExcel.csv";
   	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"CSVToExcel_out.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str(), L",");

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2:E19"))->SetIgnoreErrorOptions(IgnoreErrorType::NumberAsText);
		sheet->GetAllocatedRange()->AutoFitColumns();

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();
}
