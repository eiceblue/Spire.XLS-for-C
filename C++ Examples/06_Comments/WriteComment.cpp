#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
                wstring output_path = OUTPUTPATH;
                wstring outputFile = output_path + L"WriteComment_out.xlsx";

				//Create a workbook
				intrusive_ptr<Workbook> workbook = new Workbook();

				//Get the first worksheet
				intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

				//Creates font
				intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
				font->SetFontName(L"Arial");
				font->SetSize(11);
				font->SetKnownColor(ExcelColors::Orange);
				intrusive_ptr<ExcelFont> fontBlue = workbook->CreateExcelFont();
				fontBlue->SetKnownColor(ExcelColors::LightBlue);
				intrusive_ptr<ExcelFont> fontGreen = workbook->CreateExcelFont();
				fontGreen->SetKnownColor(ExcelColors::LightGreen);

				intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"));
				range->SetText(L"Regular comment");
				range->GetComment()->SetText(L"Regular comment");
				range->AutoFitColumns();
				//Regular comment

				range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B12"));
				range->SetText(L"Rich text comment");
				range->GetRichText()->SetFont(0, 16, font);
				range->AutoFitColumns();
				//Rich text comment
				range->GetComment()->GetRichText()->SetText(L"Rich text comment");
				range->GetComment()->GetRichText()->SetFont(0, 4, fontGreen);
				range->GetComment()->GetRichText()->SetFont(5, 9, fontBlue);

				//Save to file.
				workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
				workbook->Dispose();

}
