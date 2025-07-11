#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertHtmlStringIntoCell.xlsx";

	// Create a Workbook object
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Get the first worksheet from the workbook
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Define an HTML code to be placed in cell A1
	std::wstring htmlCode = L"<div>first line<br>second line<br>third line</div>";

	// Get the cell range for cell A1 and set the HTML string
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"));
	range->SetHtmlString(htmlCode.c_str());

	// Save the workbook to the output file in Excel 2010 format
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);

	// Dispose the workbook to release resources
	workbook->Dispose();
}

