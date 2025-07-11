#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateNestedGroup_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set the style.
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"style");
	style->GetFont()->SetColor(Spire::Xls::Color::GetCadetBlue());
	style->GetFont()->SetIsBold(true);

	//Set the summary rows appear above detail rows.
	sheet->GetPageSetup()->SetIsSummaryRowBelow(false);

	//Insert sample data to cells.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Project plan for project X");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetCellStyleName(style->GetName());

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetValue(L"Set up");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetCellStyleName(style->GetName());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetValue(L"Task 1");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetValue(L"Task 2");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4:A5"))->BorderAround(LineStyleType::Thin);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4:A5"))->BorderInside(LineStyleType::Thin);

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetValue(L"Launch");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetCellStyleName(style->GetName());
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8"))->SetValue(L"Task 1");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A9"))->SetValue(L"Task 2");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:A9"))->BorderAround(LineStyleType::Thin);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:A9"))->BorderInside(LineStyleType::Thin);

	//Group the rows that you want to group.
	sheet->GroupByRows(2, 9, false);
	sheet->GroupByRows(4, 5, false);
	sheet->GroupByRows(8, 9, false);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

