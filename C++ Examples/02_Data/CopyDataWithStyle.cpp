#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CopyDataWithStyle_result.xlsx";


	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set the values for some cells.
	intrusive_ptr<CellRange> cells = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:J50"));
	for (int i = 1; i <= 10; i++)
	{
		for (int j = 1; j <= 8; j++)
		{
			cells->Get(i, j)->SetText((std::to_wstring(i - 1) + L"," + std::to_wstring(j - 1)).c_str());
		}
	}
	//Get a source range (A1:D3).
	intrusive_ptr<CellRange> srcRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D3"));

	//Create a style object.
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"style");

	//Specify the font attribute.
	style->GetFont()->SetFontName(L"Calibri");

	//Specify the shading color.
	style->GetFont()->SetColor(Color::GetRed());

	//Specify the border attributes.
	style->GetBorders()->Get(BordersLineType::EdgeTop)->SetLineStyle(LineStyleType::Thin);
	style->GetBorders()->Get(BordersLineType::EdgeTop)->SetColor(Color::GetBlue());
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetLineStyle(LineStyleType::Thin);
	style->GetBorders()->Get(BordersLineType::EdgeBottom)->SetColor(Color::GetBlue());
	style->GetBorders()->Get(BordersLineType::EdgeLeft)->SetLineStyle(LineStyleType::Thin);
	style->GetBorders()->Get(BordersLineType::EdgeLeft)->SetColor(Color::GetBlue());
	style->GetBorders()->Get(BordersLineType::EdgeRight)->SetLineStyle(LineStyleType::Thin);
	style->GetBorders()->Get(BordersLineType::EdgeRight)->SetColor(Color::GetBlue());
	srcRange->SetCellStyleName(style->GetName());

	//Set the destination range
	intrusive_ptr<CellRange> destRange = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A12:D14"));

	//Copy the range Demo with style
	srcRange->Copy(destRange, true, true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

