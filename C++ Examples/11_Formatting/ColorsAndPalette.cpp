#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ColorsAndPalette.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Adding Orchid color to the palette at 60th index
	workbook->ChangePaletteColor(Spire::Xls::Color::GetOrchid(), 60);

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"));
	cell->SetText(L"Welcome to use Spire.XLS");

	//Set the Orchid (custom) color to the font
	cell->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetOrchid());
	cell->GetStyle()->GetFont()->SetSize(20);
	cell->AutoFitColumns();
	cell->AutoFitRows();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}