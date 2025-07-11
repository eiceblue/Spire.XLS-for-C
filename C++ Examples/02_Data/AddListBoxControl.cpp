#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddListBoxControl.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set text for cells 
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetText(L"Beijing");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8"))->SetText(L"New York");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A9"))->SetText(L"ChengDu");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A10"))->SetText(L"Paris");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetText(L"Boston");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetText(L"London");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetText(L"City :");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C13"))->GetStyle()->GetFont()->SetIsBold(true);

	//Add listbox control
	intrusive_ptr<IListBox> listBox = sheet->GetListBoxes()->AddListBox(13, 4, 120, 100);
	listBox->SetSelectionType(SelectionType::Single);
	listBox->SetSelectedIndex(2);
	listBox->SetDisplay3DShading(true);
	listBox->SetListFillRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7:A12")));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

