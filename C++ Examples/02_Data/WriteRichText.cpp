#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring fn = DATAPATH;
	wstring inputFile = fn + L"Sample.xlsx";
	wstring outputFile = output_path + L"WriteRichText_result.xlsx";

	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	intrusive_ptr<ExcelFont> fontBold = workbook->CreateExcelFont();
	fontBold->SetIsBold(true);

	intrusive_ptr<ExcelFont> fontUnderline = workbook->CreateExcelFont();
	fontUnderline->SetUnderline(FontUnderlineType::Single);

	intrusive_ptr<ExcelFont> fontItalic = workbook->CreateExcelFont();
	fontItalic->SetIsItalic(true);

	intrusive_ptr<ExcelFont> fontColor = workbook->CreateExcelFont();
	fontColor->SetKnownColor(ExcelColors::Green);

	intrusive_ptr<RichText> richText = dynamic_pointer_cast<RichText>(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"))->GetRichText());
	richText->SetText(L"Bold and underlined and italic and colored text.");
	richText->SetFont(0, 3, fontBold);
	richText->SetFont(9, 18, fontUnderline);
	richText->SetFont(24, 29, fontItalic);
	richText->SetFont(35, 41, fontColor);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

