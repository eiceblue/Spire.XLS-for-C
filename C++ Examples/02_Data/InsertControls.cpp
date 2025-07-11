#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Sample.xlsx";
	wstring outputFile = output_path + L"InsertControls_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add a textbox 
	intrusive_ptr<ITextBoxShape> textbox = sheet->GetTextBoxes()->AddTextBox(9, 2, 25, 100);
	textbox->SetText(L"Hello World");
	//Add a checkbox 
	intrusive_ptr<ICheckBox> cb = sheet->GetCheckBoxes()->AddCheckBox(11, 2, 15, 100);
	cb->SetCheckState(Spire::Xls::CheckState::Checked);
	cb->SetText(L"Check Box 1");
	//Add a RadioButton 
	intrusive_ptr<IRadioButton> rb = sheet->GetRadioButtons()->Add(13, 2, 15, 100);
	rb->SetText(L"Option 1");

	//Add a combox
	intrusive_ptr<IComboBoxShape> cbx = dynamic_pointer_cast<IComboBoxShape>(sheet->GetComboBoxes()->AddComboBox(15, 2, 15, 100));
	cbx->SetListFillRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A41:A47")));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

