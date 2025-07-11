#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FormatAColumn.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a new style
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");

	//Set the vertical alignment of the text
	style->SetVerticalAlignment(VerticalAlignType::Center);

	//Set the horizontal alignment of the text
	style->SetHorizontalAlignment(HorizontalAlignType::Center);

	//Set the font color of the text
	style->GetFont()->SetColor(Spire::Xls::Color::GetBlue());

	//Shrink the text to fit in the cell
	style->SetShrinkToFit(true);

	//Set the bottom border color of the cell to OrangeRed
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Spire::Xls::Color::GetOrangeRed());

	//Set the bottom border type of the cell to Dotted
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Dotted);

	//Apply the style to the first column
	sheet->GetColumns()->GetItem(0)->SetCellStyleName(style->GetName());

	sheet->GetColumns()->GetItem(0)->SetText(L"Test");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
