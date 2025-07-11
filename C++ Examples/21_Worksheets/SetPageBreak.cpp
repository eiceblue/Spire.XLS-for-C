#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"WorksheetSample1.xlsx";
	std::wstring outputFile = output_path + L"SetPageBreak.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set Excel Page Break Horizontally
	(dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks()))->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8")));
	(dynamic_pointer_cast<HPageBreaksCollection>(sheet->GetHPageBreaks()))->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A14")));

	//Set view mode to Preview mode
	dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->SetViewMode(ViewMode::Preview);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}