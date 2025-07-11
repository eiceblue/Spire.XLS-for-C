#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"TextBoxSampleB.xlsx";
	wstring outputFile = output + L"TextBoxWithWrapText.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the text box
	intrusive_ptr<XlsTextBoxShape> shape = dynamic_pointer_cast<XlsTextBoxShape>(sheet->GetTextBoxes()->Get(0));

	//Set wrap text
	shape->SetIsWrapText(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}