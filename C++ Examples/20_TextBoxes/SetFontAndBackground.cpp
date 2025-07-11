#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_5.xlsx";
	wstring outputFile = output + L"SetFontAndBackground.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the textbox which will be edited.
	intrusive_ptr<XlsTextBoxShape> shape = dynamic_pointer_cast<XlsTextBoxShape>(sheet->GetTextBoxes()->Get(0));

	//Set the font and background color for the textbox.
	//Set font.
	intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
	//font.IsStrikethrough = true;
	font->SetFontName(L"Century Gothic");
	font->SetSize(10);
	font->SetIsBold(true);
	font->SetColor(Spire::Xls::Color::GetBlue());
	intrusive_ptr<RichTextShape> tempVar = dynamic_pointer_cast<RichTextShape>(shape->GetRichText());
	wstring text = shape->GetText();
	tempVar->SetFont(0, text.size() - 1, font);

	//Set background color
	shape->GetFill()->SetFillType(ShapeFillType::SolidColor);
	shape->GetFill()->SetForeKnownColor(ExcelColors::BlueGray);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}