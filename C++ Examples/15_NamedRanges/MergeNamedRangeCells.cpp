#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"AllNamedRanges.xlsx";
	wstring outputFile = output_path + L"MergeNamedRangeCells.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get specific named range by index
	intrusive_ptr<INamedRange> NamedRange = workbook->GetNameRanges()->Get(0);

	//Get the range of the named range
	intrusive_ptr<IXLSRange> range = NamedRange->GetRefersToRange();

	//Merge cells
	range->Merge();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}