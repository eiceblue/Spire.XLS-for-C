#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"ProtectCell.xlsx";
	wstring outputFile = output + L"ProtectCell.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Protect cell
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->GetStyle()->SetLocked(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->GetStyle()->SetLocked(false);

	sheet->XlsWorksheetBase::Protect(L"TestPassword");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}