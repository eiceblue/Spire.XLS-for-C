#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_6.xlsx";
	wstring outputFile = output_path + L"CreateFormulaConditionalFormat.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	intrusive_ptr<XlsRange> range = sheet->GetColumns()->GetItem(0);

	//Set the conditional formatting formula and apply the rule to the chosen cell range.
	intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Add();
	xcfs->AddRange(range);
	intrusive_ptr<IConditionalFormat> conditional = xcfs->AddCondition();
	conditional->SetFormatType(ConditionalFormatType::Formula);
	conditional->SetFirstFormula(L"=($A1<$B1)");
	conditional->SetBackKnownColor(ExcelColors::Yellow);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}