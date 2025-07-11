#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Template_Xls_1.xlsx";
	wstring inputImage = inputFolder + L"Background.png";
	wstring outputFile = outputFolder + L"InsertExcelBackgroundImage_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Open an image. 
	intrusive_ptr<Stream> im = new Stream(inputFile_Img.c_str());
	intrusive_ptr<Stream> bm = Object::Convert<Stream>(im);
	//Set the image to be background image of the worksheet.
	sheet->GetPageSetup()->SetBackgoundImage(bm);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
