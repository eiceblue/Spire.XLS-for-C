#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"templateAz.xlsx";
	wstring outputFile = output_path + L"GetStyleSetStyle.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get "B4" cell
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"));
	//Get the style of cell
	intrusive_ptr<CellStyle> style = dynamic_pointer_cast<CellStyle>(range->GetStyle());
	style->GetFont()->SetFontName(L"Calibri");
	style->GetFont()->SetIsBold(true);
	style->GetFont()->SetSize(15);
	style->GetFont()->SetColor(Spire::Xls::Color::GetCornflowerBlue());

	range->SetStyle(style);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
