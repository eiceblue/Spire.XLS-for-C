#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring inputFile = fn + L"Template_Xls_1.xlsx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ApplyMultipleFontsInSingleCell_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a font object in workbook, setting the font color, size and type.
	intrusive_ptr<ExcelFont> font1 = workbook->CreateExcelFont();
	font1->SetKnownColor(ExcelColors::LightBlue);
	font1->SetIsBold(true);
	font1->SetSize(10);

	//Create another font object specifying its properties.
	intrusive_ptr<ExcelFont> font2 = workbook->CreateExcelFont();
	font2->SetKnownColor(ExcelColors::Red);
	font2->SetIsBold(true);
	font2->SetIsItalic(true);
	font2->SetFontName(L"Times New Roman");
	font2->SetSize(11);

	//Write a RichText string to the cell 'A1', and set the font for it.
	intrusive_ptr<RichText> richText = dynamic_pointer_cast<RichText>(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"H5"))->GetRichText());
	richText->SetText(L"This document was created with Spire.XLS for .NET.");
	richText->SetFont(0, 29, font1);
	richText->SetFont(31, 48, font2);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

