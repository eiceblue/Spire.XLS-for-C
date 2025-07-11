#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"FontStyles.xlsx";
	wstring outputFile = output_path + L"FontStyles.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set font style
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetFontName(L"Comic Sans MS");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D2"))->GetStyle()->GetFont()->SetFontName(L"Corbel");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:D7"))->GetStyle()->GetFont()->SetFontName(L"Aleo");

	//Set font size
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetSize(45);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D3"))->GetStyle()->GetFont()->SetSize(25);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:D7"))->GetStyle()->GetFont()->SetSize(12);

	//Set excel cell data to be bold
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D2"))->GetStyle()->GetFont()->SetIsBold(true);

	//Set excel cell data to be underline
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:B7"))->GetStyle()->GetFont()->SetUnderline(FontUnderlineType::Single);

	//set excel cell data color
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetCornflowerBlue());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2:D2"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetCadetBlue());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:D7"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetFirebrick());

	//set excel cell data to be italic
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3:D7"))->GetStyle()->GetFont()->SetIsItalic(true);

	//Add strikethrough
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D3"))->GetStyle()->GetFont()->SetIsStrikethrough(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D7"))->GetStyle()->GetFont()->SetIsStrikethrough(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}