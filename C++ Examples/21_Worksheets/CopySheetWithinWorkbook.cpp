#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"Template_Xls_4.xlsx";
	std::wstring outputFile = output_path + L"CopySheetWithinWorkbook.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first and the second worksheets.
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Add(L"MySheet"));
	intrusive_ptr<CellRange> sourceRange = dynamic_pointer_cast<CellRange>(sheet->GetAllocatedRange());

	//Copy the first worksheet to the second one.
	sheet->Copy(sourceRange, sheet1, sheet->GetFirstRow(), sheet->GetFirstColumn(), true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

