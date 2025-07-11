#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"SetDBNumFormatting_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->CreateEmptySheets(1);

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set value for cells
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetNumberValue(123);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetNumberValue(456);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetNumberValue(789);

	//Get the cell range
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:A3"));

	//Set the DB num format
	range->SetNumberFormat(L"[DBNum2][$-804]General");

	//Auto fit columns
	range->AutoFitColumns();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
