#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"NamedRanges.xlsx";
	wstring outputFile = output_path + L"NamedRanges.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Creating a named range
	intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Add(L"NewNamedRange");
	//Setting the range of the named range
	NamedRange->SetRefersToRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A8:E12")));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}