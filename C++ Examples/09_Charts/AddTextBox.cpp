#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"AddTextBox.xlsx";
	wstring outputFile = output_path + L"AddTextBox.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the first chart
	intrusive_ptr<Chart> chart = dynamic_pointer_cast<Chart>(sheet->GetCharts()->Get(0));

	//Add a Textbox
	intrusive_ptr<ITextBoxLinkShape> textbox = chart->GetShapes()->AddTextBox();
	textbox->SetWidth(1200);
	textbox->SetHeight(320);
	textbox->SetLeft(1000);
	textbox->SetTop(480);
	textbox->SetText(L"This is a textbox");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
