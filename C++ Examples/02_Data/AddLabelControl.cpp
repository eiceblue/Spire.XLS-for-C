#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring inputFile = fn + L"Sample.xlsx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddLabelControl.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add a label control
	intrusive_ptr<ILabelShape> label = sheet->GetLabelShapes()->AddLabel(10, 2, 30, 200);
	label->SetText(L"This is a Label Control");
	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

