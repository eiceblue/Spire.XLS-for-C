#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SetHeaderFooter.xlsx";
	wstring outputFile = output_path + L"SetHeaderFooter.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set left header,"Arial Unicode MS" is font name, L"18" is font size.
	sheet->GetPageSetup()->SetLeftHeader(L"&\"Arial Unicode MS\"&14 Spire.XLS for C++ ");

	//Set center footer 
	sheet->GetPageSetup()->SetCenterFooter(L"Footer Text");

	sheet->SetViewMode(ViewMode::Layout);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}