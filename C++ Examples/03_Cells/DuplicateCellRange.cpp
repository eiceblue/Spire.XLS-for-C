#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_1.xlsx";
	wstring outputFile = output_path + L"DuplicateCellRange_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Copy data from source range to destination range and maintain the format.
	sheet->Copy(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6:F6")), dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A16:F16")), true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
