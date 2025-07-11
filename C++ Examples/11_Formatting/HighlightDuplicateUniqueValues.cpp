#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_6.xlsx";
	wstring outputFile = output_path + L"HighlightDuplicateUniqueValues.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
	intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C10")));
	intrusive_ptr<IConditionalFormat> format1 = xcfs->AddCondition();
	format1->SetFormatType(ConditionalFormatType::DuplicateValues);
	format1->SetBackColor(Spire::Xls::Color::GetIndianRed());

	//Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
	intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
	xcfs1->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:C10")));
	intrusive_ptr<IConditionalFormat> format2 = xcfs->AddCondition();
	format2->SetFormatType(ConditionalFormatType::UniqueValues);
	format2->SetBackColor(Spire::Xls::Color::GetYellow());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}