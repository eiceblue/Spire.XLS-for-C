#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"CommonTemplate1.xlsx";
	wstring outputFile = output_path + L"AddHyperlinkToText.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add url link
	intrusive_ptr<HyperLink> UrlLink = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D10"))));
	UrlLink->SetTextToDisplay(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D10"))->GetText());
	UrlLink->SetType(HyperLinkType::Url);
	UrlLink->SetAddress(L"http://en.wikipedia.org/wiki/Chicago");

	//Add email link
	intrusive_ptr<XlsHyperLink> MailLink = sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E10")));
	MailLink->SetTextToDisplay(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E10"))->GetText());
	MailLink->SetType(HyperLinkType::Url);
	MailLink->SetAddress(L"mailto:Amor.Aqua@gmail.com");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}