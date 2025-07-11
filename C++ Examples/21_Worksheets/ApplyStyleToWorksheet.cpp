#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"WorksheetSample1.xlsx";
	std::wstring outputFile = output_path + L"ApplyStyleToWorksheet.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create a cell style
	intrusive_ptr<CellStyle> style = workbook->GetStyles()->Add(L"newStyle");
	style->SetColor(Spire::Xls::Color::GetLightBlue());
	style->GetFont()->SetColor(Spire::Xls::Color::GetWhite());
	style->GetFont()->SetSize(15);
	style->GetFont()->SetIsBold(true);
	//Apply the style to the first worksheet
	sheet->ApplyStyle(style);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
