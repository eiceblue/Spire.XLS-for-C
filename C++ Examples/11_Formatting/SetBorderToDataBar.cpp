#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"Template_Xls_9.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create an intrusive pointer to a Workbook object
	intrusive_ptr<Workbook>workbook = new Workbook();

	// Load the workbook from the input file
	workbook->LoadFromFile(inputFile.c_str());

	// Get the first worksheet from the workbook
	intrusive_ptr<Worksheet>sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Get the conditional formats from the worksheet and retrieve the first format
	intrusive_ptr<XlsConditionalFormats> xcfs = sheet->GetConditionalFormats()->Get(0); intrusive_ptr<IConditionalFormat>cf = xcfs->Get(0);
	
	// Get the data bar from the first format and set its border type and color
	intrusive_ptr<Spire::Xls::DataBar> dataBar1 = cf->GetDataBar();
	dataBar1->GetBarBorder()->SetType(DataBarBorderType::DataBarBorderSolid);
	dataBar1->GetBarBorder()->SetColor(Color::GetRed());

	// Set a numerical value in cell E1 of the worksheet
	sheet->GetRange(L"E1")->SetNumberValue(200);

	// Add a new conditional format to the worksheet and specify the range as E1
	intrusive_ptr<XlsConditionalFormats>xcfs2 = sheet->GetConditionalFormats()->Add(); xcfs2->AddRange(sheet->GetRange(L"E1"));
	
	// Add a condition to the new format and set its type, bar border type, color, and bar color
	intrusive_ptr<IConditionalFormat>cf2 = xcfs2->AddCondition(); cf2->SetFormatType(ConditionalFormatType::DataBar); cf2->GetDataBar()->GetBarBorder()->SetType(DataBarBorderType::DataBarBorderSolid); cf2->GetDataBar()->GetBarBorder()->SetColor(Color::GetRed()); cf2->GetDataBar()->SetBarColor(Color::GetGreenYellow());
	
	// Save the workbook to the output file in Excel 2013 format
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);

	// Dispose of the workbook object
	workbook->Dispose();
}

