#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ApplyConditionalFormatting.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Insert data to cell range from A1 to C4.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetNumberValue(582);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetNumberValue(234);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetNumberValue(314);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetNumberValue(50);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->SetNumberValue(150);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B2"))->SetNumberValue(894);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B3"))->SetNumberValue(560);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B4"))->SetNumberValue(900);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetNumberValue(134);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetNumberValue(700);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C3"))->SetNumberValue(920);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C4"))->SetNumberValue(450);
	sheet->GetAllocatedRange()->SetRowHeight(15);
	sheet->GetAllocatedRange()->SetColumnWidth(17);

	//Create conditional formatting rule.
	intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
	xcfs1->AddRange(sheet->GetAllocatedRange());
	intrusive_ptr<IConditionalFormat> format1 = xcfs1->AddCondition();
	format1->SetFormatType(ConditionalFormatType::CellValue);
	format1->SetFirstFormula(L"800");
	format1->SetOperator(ComparisonOperatorType::Greater);
	format1->SetFontColor(Spire::Xls::Color::GetRed());
	format1->SetBackColor(Spire::Xls::Color::GetLightSalmon());

	//Create conditional formatting rule.
	intrusive_ptr<XlsConditionalFormats> xcfs2 = sheet->GetConditionalFormats()->Add();
	xcfs2->AddRange(sheet->GetAllocatedRange());
	intrusive_ptr<IConditionalFormat> format2 = xcfs1->AddCondition();
	format2->SetFormatType(ConditionalFormatType::CellValue);
	format2->SetFirstFormula(L"300");
	format2->SetOperator(ComparisonOperatorType::Less);
	format2->SetFontColor(Spire::Xls::Color::GetGreen());
	format2->SetBackColor(Spire::Xls::Color::GetLightBlue());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
