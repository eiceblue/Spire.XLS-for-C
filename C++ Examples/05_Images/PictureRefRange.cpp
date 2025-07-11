#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"PictureRefRange.xlsx";
	wstring outputFile = outputFolder + L"PictureRefRange_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetValue(L"Spire.XLS");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetValue(L"E-iceblue");

	//Get the first picture in worksheet
	intrusive_ptr<XlsBitmapShape> picture = sheet->GetPictures()->Get(0);

	//Set the reference range of the picture to A1:B3
	picture->SetRefRange(L"A1:B3");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
