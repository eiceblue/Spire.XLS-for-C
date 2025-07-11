#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LinkToExternalFile.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"));

	//Add hyperlink in the range
	intrusive_ptr<HyperLink> hyperlink = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(range));

	//Set the link type
	hyperlink->SetType(HyperLinkType::Workbook);

	//Set the display text
	hyperlink->SetTextToDisplay(L"Link to Sheet2 cell C5");

	//Set the address
	hyperlink->SetAddress(L"Sheet2!C5");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}