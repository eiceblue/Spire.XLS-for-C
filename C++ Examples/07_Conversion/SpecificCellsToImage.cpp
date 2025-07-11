#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConversionSample1.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"Image";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Specify Cell Ranges and Save to certain Image formats
	sheet->ToImage(1, 1, 7, 5)->Save((outputFile + L"SpecificCellsToImage.png").c_str());
	sheet->ToImage(8, 1, 15, 5)->Save((outputFile + L"SpecificCellsToImage.jpg").c_str());
	sheet->ToImage(17, 1, 23, 5)->Save((outputFile + L"SpecificCellsToImage.bmp").c_str());

	workbook->Dispose();
}
