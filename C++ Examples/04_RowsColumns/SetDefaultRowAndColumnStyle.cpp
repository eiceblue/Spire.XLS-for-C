#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"SetDefaultRowAndColumnStyle_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a cell style and set the color
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"Mystyle");
	style->SetColor(Spire::Xls::Color::GetYellow());

	//Set the default style for the first row and column 
	sheet->SetDefaultRowStyle(1, style);
	sheet->SetDefaultColumnStyle(1, style);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
