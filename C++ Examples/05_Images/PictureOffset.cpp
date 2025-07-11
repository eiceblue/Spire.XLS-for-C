#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Logo.png";
	wstring outputFile = outputFolder + L"PictureOffset_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Insert a picture
	intrusive_ptr<ExcelPicture> pic = ExcelPicture::Dynamic_cast<ExcelPicture>(sheet->GetPictures()->Add(2, 2, inputFile.c_str()));

	//Set left offset and top offset from the current range
	pic->SetLeftColumnOffset(200);
	pic->SetTopRowOffset(100);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
