#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"WorksheetSample1.xlsx";
	std::wstring outputFile = output_path + L"SetMargins.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set margins for top, bottom, left and right, here the unit of measure is Inch
	sheet->GetPageSetup()->SetTopMargin(0.3);
	sheet->GetPageSetup()->SetBottomMargin(1);
	sheet->GetPageSetup()->SetLeftMargin(0.2);
	sheet->GetPageSetup()->SetRightMargin(1);
	//Set the header margin and footer margin
	sheet->GetPageSetup()->SetHeaderMarginInch(0.1);
	sheet->GetPageSetup()->SetFooterMarginInch(0.5);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}