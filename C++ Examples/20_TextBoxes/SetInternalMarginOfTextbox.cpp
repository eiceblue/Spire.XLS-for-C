#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_4.xlsx";
	wstring outputFile = output + L"SetInternalMarginOfTextbox.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add a textbox to the sheet and set its position and size.
	intrusive_ptr<XlsTextBoxShape> textbox = dynamic_pointer_cast<XlsTextBoxShape>(sheet->GetTextBoxes()->AddTextBox(4, 2, 100, 300));

	//Set the text on the textbox.
	textbox->SetText(L"Insert TextBox in Excel and set the margin for the text");
	textbox->SetHAlignment(CommentHAlignType::Center);
	textbox->SetVAlignment(CommentVAlignType::Center);

	//Set the inner margins of the contents.
	textbox->SetInnerLeftMargin(1);
	textbox->SetInnerRightMargin(3);
	textbox->SetInnerTopMargin(1);
	textbox->SetInnerBottomMargin(1);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}