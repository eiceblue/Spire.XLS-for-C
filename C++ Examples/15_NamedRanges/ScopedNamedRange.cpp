#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"CreateTable.xlsx";
	wstring outputFile = output_path + L"ScopedNamedRange.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add range name
	intrusive_ptr<INamedRange> namedRange = sheet->GetNames()->Add(L"Range1");

	//Define the range
	namedRange->SetRefersToRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D19")));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}