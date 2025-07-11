#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ForegroundAndBackground.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->SetVersion(ExcelVersion::Version2010);

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a new style
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle1");

	//Set filling pattern type
	style->GetInterior()->SetFillPattern(ExcelPatternType::Gradient);

	//Set filling Background color
	style->GetInterior()->GetGradient()->SetBackKnownColor(ExcelColors::Green);

	//Set filling Foreground color
	style->GetInterior()->GetGradient()->SetForeKnownColor(ExcelColors::Yellow);

	style->GetInterior()->GetGradient()->SetGradientStyle(GradientStyleType::From_Center);

	//Apply the style to  "B2" cell
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetCellStyleName(style->GetName());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetText(L"Test");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetRowHeight(30);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetColumnWidth(50);


	//Create a new style
	style = workbook->GetStyles()->Add(L"newStyle2");

	//Set filling pattern type
	style->GetInterior()->SetFillPattern(ExcelPatternType::Gradient);

	//Set filling Foreground color
	style->GetInterior()->GetGradient()->SetForeKnownColor(ExcelColors::Red);

	//Apply the style to  "B4" cell
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetCellStyleName(style->GetName());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetRowHeight(30);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetColumnWidth(60);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}