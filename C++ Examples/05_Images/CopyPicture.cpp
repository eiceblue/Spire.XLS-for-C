#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"ReadImages.xlsx";
	wstring outputFile = outputFolder + L"CopyPicture_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add a new worksheet as destination sheet
	intrusive_ptr<Worksheet> destinationSheet = workbook->GetWorksheets()->Add(L"DestSheet");
	//Get the first picture from the first worksheet
	intrusive_ptr<XlsBitmapShape> sourcePicture = sheet->GetPictures()->Get(0);
	//Get the image
	intrusive_ptr<Stream> image = sourcePicture->GetPicture();
	//Add the image into the added worksheet 
	destinationSheet->GetPictures()->Add(2, 2, image);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
