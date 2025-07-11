#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"UseExplicitLineBreaks_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet1 = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Specify a cell range
	intrusive_ptr<CellRange> c5 = dynamic_pointer_cast<CellRange>(sheet1->GetRange(L"C5"));

	//Set the cell width for specified range
	sheet1->SetColumnWidth(c5->GetColumn(), 70);

	//Put the string value with explicit line breaks
	c5->SetValue(L"Spire.XLS for C++ is a professional Excel C++ API\n that can be used to create, read, write and convert Excel files in any type of C++ application.\n Spire.XLS for C++ offers object model Excel API for speeding up Excel programming\n in C++ platform -create new Excel documents from template, edit existing Excel documents and convert Excel files.");

	//Set Text wrap
	c5->SetIsWrapText(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
