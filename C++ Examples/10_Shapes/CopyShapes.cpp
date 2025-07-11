#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CopyShapes.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create line shape
	auto line = sheet->GetTypedLines()->AddLine();
	line->SetTop(50);
	line->SetLeft(30);
	line->SetWidth(30);
	line->SetHeight(50);
	line->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrowDiamond);
	line->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);

	intrusive_ptr<Worksheet> CopyShapes = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));
	//Copy the line into other sheet
	CopyShapes->GetTypedLines()->AddCopy(line);

	//Create a button and then copy into other sheet
	auto button = sheet->GetTypedRadioButtons()->Add(5, 5, 20, 20);
	CopyShapes->GetTypedRadioButtons()->AddCopy(button);

	//Create a textbox and then copy into other sheet
	auto textbox = sheet->GetTypedTextBoxes()->AddTextBox(5, 7, 50, 100);
	CopyShapes->GetTypedTextBoxes()->AddCopy(textbox);

	//Create a checkbox and then copy into other sheet
	auto checkbox = sheet->GetTypedCheckBoxes()->AddCheckBox(10, 1, 20, 20);
	CopyShapes->GetTypedCheckBoxes()->AddCopy(checkbox);

	//Create a comboboxes and then copy into other sheet
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A14"))->SetValue(L"1");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A15"))->SetValue(L"2");
	auto ComboBoxes = sheet->GetTypedComboBoxes()->AddComboBox(10, 5, 30, 30);
	ComboBoxes->SetListFillRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A14:A15")));
	CopyShapes->GetTypedComboBoxes()->AddCopy(ComboBoxes);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
