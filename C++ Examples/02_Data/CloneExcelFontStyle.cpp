#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CloneExcelFontStyle_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add the text to the Excel sheet cell range A1.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Text1");

	//Set A1 cell range's CellStyle.
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"style");
	style->GetFont()->SetFontName(L"Calibri");
	style->GetFont()->SetColor(Spire::Xls::Color::GetRed());
	style->GetFont()->SetSize(12);
	style->GetFont()->SetIsBold(true);
	style->GetFont()->SetIsItalic(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetCellStyleName(style->GetName());

	//Clone the same style for B2 cell GetRange.
	intrusive_ptr<CellStyle> csOrieign = style->clone();
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetText(L"Text2");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetCellStyleName(csOrieign->GetName());

	//Clone the same style for C3 cell GetRange and then reset the font color for the text.
	intrusive_ptr<CellStyle> csGreen = style->clone();
	csGreen->GetFont()->SetColor(Spire::Xls::Color::GetGreen());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetText(L"Text3");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetCellStyleName(csGreen->GetName());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

