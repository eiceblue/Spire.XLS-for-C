#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SampleB_2.xlsx";
	wstring outputFile = output_path + L"FormatCellsWithStyle.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Create a style
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");
	//Set the shading color
	style->SetColor(Spire::Xls::Color::GetDarkGray());
	//Set the font color
	style->GetFont()->SetColor(Spire::Xls::Color::GetWhite());
	//Set font name
	style->GetFont()->SetFontName(L"Times New Roman");
	//Set font size
	style->GetFont()->SetSize(12);
	//Set bold for the font
	style->GetFont()->SetIsBold(true);
	//Set text rotation
	style->SetRotation(45);
	//Set alignment
	style->SetHorizontalAlignment(HorizontalAlignType::Center);
	style->SetVerticalAlignment(VerticalAlignType::Center);

	//Set the style for the specific range
	dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetRange(L"A1:J1")->SetCellStyleName(style->GetName());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
