#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	wstring outputFile = output_path + L"SetFormulaWithNamedRange.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a named range
	intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Add(L"MyNamedRange");
	//Refers to range
	NamedRange->SetRefersToRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10:B12")));

	//Set the formula of range to named range
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13"))->SetFormula(L"=SUM(MyNamedRange)");

	//Set value of ranges
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))->SetNumberValue(10);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"))->SetNumberValue(20);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B12"))->SetNumberValue(30);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}