#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UsingStyleObject.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = workbook->GetWorksheets()->Add(L"new sheet");

	//Access the "B1" cell from the worksheet
	intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"));

	//Add some value to the "B1" cell
	cell->SetText(L"Hello Spire!");

	//Create a new style
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");

	//Set the vertical alignment of the text in the "B1" cell
	style->SetVerticalAlignment(VerticalAlignType::Center);

	//Set the horizontal alignment of the text in the "B1" cell
	style->SetHorizontalAlignment(HorizontalAlignType::Center);

	//Set the font color of the text in the "B1" cell
	style->GetFont()->SetColor(Spire::Xls::Color::GetBlue());

	//Shrink the text to fit in the cell
	style->SetShrinkToFit(true);

	//Set the bottom border color of the cell to GreenYellow
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Xls::Color::GetGreenYellow());

	//Set the bottom border type of the cell to Medium
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Medium);

	//Assign the Style object to the "B1" cell
	cell->SetStyle(style);


	//Apply the same style to some other cells
	sheet->GetRange(L"B4")->SetStyle(style);
	sheet->GetRange(L"B4")->SetText(L"Test");
	sheet->GetRange(L"C3")->SetCellStyleName(style->GetName());
	sheet->GetRange(L"C3")->SetText(L"Welcome to use Spire.XLS");
	sheet->GetRange(L"D4")->SetStyle(style);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

}