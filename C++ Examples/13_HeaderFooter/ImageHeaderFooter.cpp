#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ImageHeaderFooter.xlsx";
	wstring outputFile = output_path + L"ImageHeaderFooter.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Load an image from disk
	intrusive_ptr<Stream> image = new Stream(DATAPATH L"/Demo/Logo.png");

	//Set the image header
	sheet->GetPageSetup()->SetLeftHeaderImage(image);
	sheet->GetPageSetup()->SetLeftHeader(L"&G");

	//Set the image footer
	sheet->GetPageSetup()->SetCenterFooterImage(image);
	sheet->GetPageSetup()->SetCenterFooter(L"&G");

	//Set the view mode of the sheet
	sheet->SetViewMode(ViewMode::Layout);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}