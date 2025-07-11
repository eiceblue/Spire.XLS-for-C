#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"WorksheetSample2.xlsx";
	std::wstring outputFile = output_path + L"ShowOrHideGridLine.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first and second worksheet
	intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	intrusive_ptr<Worksheet> sheet2 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));

	//Hide grid line in the first worksheet
	sheet1->SetGridLinesVisible(false);
	//Show grid line in the first worksheet
	sheet2->SetGridLinesVisible(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}