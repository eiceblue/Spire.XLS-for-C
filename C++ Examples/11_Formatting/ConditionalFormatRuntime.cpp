#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ConditionalFormatRuntime.xlsx";
	wstring outputFile = output_path + L"ConditionalFormatRuntime.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create conditional formatting rule
	intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
	xcfs1->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1:D1")));
	intrusive_ptr<IConditionalFormat> cf1 = xcfs1->AddCondition();
	cf1->SetFormatType(ConditionalFormatType::CellValue);
	cf1->SetFirstFormula(L"150");
	cf1->SetOperator(ComparisonOperatorType::Greater);
	cf1->SetFontColor(Spire::Xls::Color::GetRed());
	cf1->SetBackColor(Spire::Xls::Color::GetLightBlue());

	intrusive_ptr<XlsConditionalFormats> xcfs2 = sheet->GetConditionalFormats()->Add();
	xcfs2->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A2:D2")));
	intrusive_ptr<IConditionalFormat> cf2 = xcfs2->AddCondition();
	cf2->SetFormatType(ConditionalFormatType::CellValue);
	cf2->SetFirstFormula(L"500");
	cf2->SetOperator(ComparisonOperatorType::Less);
	//Set border color
	cf2->SetLeftBorderColor(Spire::Xls::Color::GetPink());
	cf2->SetRightBorderColor(Spire::Xls::Color::GetPink());
	cf2->SetTopBorderColor(Spire::Xls::Color::GetDeepSkyBlue());
	cf2->SetBottomBorderColor(Spire::Xls::Color::GetDeepSkyBlue());
	cf2->SetLeftBorderStyle(LineStyleType::Medium);
	cf2->SetRightBorderStyle(LineStyleType::Thick);
	cf2->SetTopBorderStyle(LineStyleType::Double);
	cf2->SetBottomBorderStyle(LineStyleType::Double);

	//Create conditional formatting rule
	intrusive_ptr<XlsConditionalFormats> xcfs3 = sheet->GetConditionalFormats()->Add();
	xcfs3->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A3:D3")));
	intrusive_ptr<IConditionalFormat> cf3 = xcfs3->AddCondition();
	cf3->SetFormatType(ConditionalFormatType::CellValue);
	cf3->SetFirstFormula(L"300");
	cf3->SetSecondFormula(L"500");
	cf3->SetOperator(ComparisonOperatorType::Between);
	cf3->SetBackColor(Spire::Xls::Color::GetYellow());

	//Create conditional formatting rule
	intrusive_ptr<XlsConditionalFormats> xcfs4 = sheet->GetConditionalFormats()->Add();
	xcfs4->AddRange(dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A4:D4")));
	intrusive_ptr<IConditionalFormat> cf4 = xcfs4->AddCondition();
	cf4->SetFormatType(ConditionalFormatType::CellValue);
	cf4->SetFirstFormula(L"100");
	cf4->SetSecondFormula(L"200");
	cf4->SetOperator(ComparisonOperatorType::NotBetween);
	//Set fill pattern type
	cf4->SetFillPattern(ExcelPatternType::ReverseDiagonalStripe);
	//Set foreground color
	cf4->SetColor(Spire::Xls::Color::FromArgb(255, 255, 0));
	//Set background color
	cf4->SetBackColor(Spire::Xls::Color::FromArgb(0, 255, 255));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}