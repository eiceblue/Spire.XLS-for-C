#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UsePredefinedStyles.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a new style
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");
	style->GetFont()->SetFontName(L"Calibri");
	style->GetFont()->SetIsBold(true);
	style->GetFont()->SetSize(15);
	style->GetFont()->SetColor(Spire::Xls::Color::GetCornflowerBlue());

	//Get "B5" cell
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B5"));
	range->SetText(L"Welcome to use Spire.XLS");
	range->SetCellStyleName(style->GetName());
	range->AutoFitColumns();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}