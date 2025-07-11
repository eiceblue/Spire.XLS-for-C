#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetTrafficLightsIcons.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add some data to the Excel sheet cell range and set the format for them.
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Traffic Lights");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetNumberValue(0.95);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2"))->SetNumberFormat(L"0%");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetNumberValue(0.5);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3"))->SetNumberFormat(L"0%");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetNumberValue(0.1);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4"))->SetNumberFormat(L"0%");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetNumberValue(0.9);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A5"))->SetNumberFormat(L"0%");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6"))->SetNumberValue(0.7);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A6"))->SetNumberFormat(L"0%");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetNumberValue(0.6);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A7"))->SetNumberFormat(L"0%");

	//Set the height of row and width of column for Excel cell range.
	sheet->GetAllocatedRange()->SetRowHeight(20);
	sheet->GetAllocatedRange()->SetColumnWidth(25);

	//Add a conditional formatting.
	intrusive_ptr<XlsConditionalFormats> conditional = sheet->GetConditionalFormats()->Add();
	conditional->AddRange(sheet->GetAllocatedRange());
	intrusive_ptr<IConditionalFormat> format1 = conditional->AddCondition();

	//Add a conditional formatting of cell range and set its type to CellValue.
	format1->SetFormatType(ConditionalFormatType::CellValue);
	format1->SetFirstFormula(L"300");
	format1->SetOperator(ComparisonOperatorType::Less);
	format1->SetFontColor(Spire::Xls::Color::GetBlack());
	format1->SetBackColor(Spire::Xls::Color::GetLightSkyBlue());

	//Add a conditional formatting of cell range and set its type to IconSet.
	conditional->AddRange(sheet->GetAllocatedRange());
	intrusive_ptr<IConditionalFormat> format = conditional->AddCondition();
	format->SetFormatType(ConditionalFormatType::IconSet);
	format->GetIconSet()->SetIconSetType(IconSetType::ThreeTrafficLights1);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}