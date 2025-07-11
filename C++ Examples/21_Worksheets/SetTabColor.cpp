#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"SetTabColor.xlsx";
	std::wstring outputFile = output_path + L"SetTabColor.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Set the tab color of first sheet to be red 
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	sheet->SetTabColor(Spire::Xls::Color::GetRed());

	//Set the tab color of first sheet to be green 
	sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1));
	sheet->SetTabColor(Spire::Xls::Color::GetGreen());

	//Set the tab color of first sheet to be blue 
	sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(2));
	sheet->SetTabColor(Spire::Xls::Color::GetLightBlue());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}