#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"CommonTemplate.xlsx";
	wstring outputFile = outputFolder + L"SetCellFillPattern_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set cell color
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B7:F7"))->GetStyle()->SetColor(Spire::Xls::Color::GetYellow());
	//Set cell fill pattern
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8:F8"))->GetStyle()->SetFillPattern(ExcelPatternType::Percent125Gray);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
