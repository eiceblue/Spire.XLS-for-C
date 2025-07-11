#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_4.xlsx";
	wstring outputFile = output + L"AddPageBreakInXlsFile.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add page break in Excel file.
	(dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks()))->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E4")));
	(dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks()))->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4")));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}