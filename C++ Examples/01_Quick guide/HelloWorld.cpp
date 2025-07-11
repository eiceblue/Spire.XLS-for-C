#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring output_path = OUTPUTPATH;
	std::wstring outputFile = output_path + L"HelloWorld.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	//Get the first sheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Hello World");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->AutoFitColumns();

	//Save to file
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);
	workbook->Dispose();
}

