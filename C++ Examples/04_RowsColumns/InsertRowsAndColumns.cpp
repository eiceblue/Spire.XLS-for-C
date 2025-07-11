#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"InsertRowsAndColumns.xls";
	wstring outputFile = outputFolder + L"InsertRowsAndColumns_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Inserting a row into the worksheet 
	sheet->InsertRow(2);
	//Inserting a column into the worksheet 
	sheet->InsertColumn(2);
	//Inserting multiple rows into the worksheet
	sheet->InsertRow(5, 2);
	//Inserting multiple columns into the worksheet
	sheet->InsertColumn(5, 2);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
