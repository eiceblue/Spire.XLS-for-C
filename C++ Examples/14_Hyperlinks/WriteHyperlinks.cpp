#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"WriteHyperlinks.xlsx";
	wstring outputFile = output_path + L"WriteHyperlinks.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B9"))->SetText(L"Home page");
	intrusive_ptr<HyperLink> hylink1 = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B10"))));
	hylink1->SetType(HyperLinkType::Url);
	hylink1->SetAddress(L"(http://www.e-iceblue.com)");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B11"))->SetText(L"Support");
	intrusive_ptr<HyperLink> hylink2 = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B12"))));
	hylink2->SetType(HyperLinkType::Url);
	hylink2->SetAddress(L"mailto:support@e-iceblue.com");

	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B13"))->SetText(L"Forum");
	intrusive_ptr<HyperLink> hylink3 = dynamic_pointer_cast<HyperLink>(sheet->GetHyperLinks()->Add(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B14"))));
	hylink3->SetType(HyperLinkType::Url);
	hylink3->SetAddress(L"https://www.e-iceblue.com/forum/");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}