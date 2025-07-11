#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ApplySubscriptAndSuperscript_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetText(L"This is an example of Subscript:");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2"))->SetText(L"This is an example of Superscript:");

	//Set the rtf value of L"B3" to L"R100-0.06".
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"));
	range->GetRichText()->SetText(L"R100-0.06");

	//Create a font. Set the IsSubscript property of the font to L"true".
	intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
	font->SetIsSubscript(true);
	font->SetColor(Spire::Xls::Color::GetGreen());

	//Set font for specified range of the text in L"B3".
	range->GetRichText()->SetFont(4, 8, font);

	//Set the rtf value of L"D3" to L"a2 + b2 = c2".
	range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D3"));
	range->GetRichText()->SetText(L"a2 + b2 = c2");

	//Create a font. Set the IsSuperscript property of the font to L"true".
	font = workbook->CreateExcelFont();
	font->SetIsSuperscript(true);

	//Set font for specified range of the text in L"D3".
	range->GetRichText()->SetFont(1, 1, font);
	range->GetRichText()->SetFont(6, 6, font);
	range->GetRichText()->SetFont(11, 11, font);

	sheet->GetAllocatedRange()->AutoFitColumns();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

