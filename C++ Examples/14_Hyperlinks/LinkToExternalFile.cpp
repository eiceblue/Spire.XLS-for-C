#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LinkToExternalFile.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(1, 1));

	//Add hyperlink in the range
	intrusive_ptr<HyperLink> hyperlink = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(range));

	//Set the link type
	hyperlink->SetType(HyperLinkType::File);

	//Set the display text
	hyperlink->SetTextToDisplay(L"Link To External File");

	//Set file address
	hyperlink->SetAddress((input_path + L"SampeB_4.xlsx").c_str());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}