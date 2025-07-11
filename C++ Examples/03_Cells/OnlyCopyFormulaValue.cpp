#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"CopyOnlyFormulaValue1.xlsx";
	wstring outputFile = outputFolder + L"CopyOnlyFormulaValue_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set the copy option
	CopyRangeOptions copyOptions = CopyRangeOptions::OnlyCopyFormulaValue;

	intrusive_ptr<CellRange> sourceRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6:E6"));
	sheet->Copy(sourceRange, dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:E8")), copyOptions);

	sourceRange->Copy(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10:E10")), copyOptions);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
