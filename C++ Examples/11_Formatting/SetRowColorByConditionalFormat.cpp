#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_4.xlsx";
	wstring outputFile = output_path + L"SetRowColorByConditionalFormat.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Select the range that you want to format.
	intrusive_ptr<CellRange> dataRange = dynamic_pointer_cast<CellRange>(sheet->GetAllocatedRange());

	//Set conditional formatting.
	intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(dataRange);
	intrusive_ptr<IConditionalFormat> format1 = xcfs->AddCondition();
	//Determines the cells to format.
	format1->SetFirstFormula(L"=MOD(ROW(),2)=0");
	//Set conditional formatting type
	format1->SetFormatType(ConditionalFormatType::Formula);
	//Set the color.
	format1->SetBackColor(Spire::Xls::Color::GetLightSeaGreen());

	//Set the backcolor of the odd rows as Yellow.
	intrusive_ptr<XlsConditionalFormats> xcfs1 = sheet->GetConditionalFormats()->Add();
	xcfs1->AddRange(dataRange);
	intrusive_ptr<IConditionalFormat> format2 = xcfs->AddCondition();
	format2->SetFirstFormula(L"=MOD(ROW(),2)=1");
	format2->SetFormatType(ConditionalFormatType::Formula);
	format2->SetBackColor(Spire::Xls::Color::GetYellow());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}