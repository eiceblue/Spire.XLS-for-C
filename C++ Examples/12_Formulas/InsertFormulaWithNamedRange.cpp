#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertFormulaWithNamedRange.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set value
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"1");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetValue(L"1");

	//Create a named range
	intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Add(L"NewNamedRange");

	NamedRange->SetNameLocal(L"=SUM(A1+A2)");

	//Set the formula
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetFormula(L"NewNamedRange");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}