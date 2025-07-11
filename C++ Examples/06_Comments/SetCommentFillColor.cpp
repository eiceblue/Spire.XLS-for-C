#include "pch.h"
using namespace Spire::Xls;

int main() {
    wstring output_path = OUTPUTPATH;
    wstring outputFile = output_path + L"SetCommentFillColor.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create Excel font
	intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
	font->SetFontName(L"Arial");
	font->SetSize(11);
	font->SetKnownColor(ExcelColors::Orange);

	//Add the comment
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"));
	range->GetComment()->SetText(L"This is a comment");
	wstring text = range->GetComment()->GetText();
	range->GetComment()->GetRichText()->SetFont(0, (text.size() - 1), font);

	//Set comment Color
	range->GetComment()->GetFill()->SetFillType(ShapeFillType::SolidColor);
	range->GetComment()->GetFill()->SetForeColor(Spire::Xls::Color::GetSkyBlue());

	range->GetComment()->SetVisible(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
