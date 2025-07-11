#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"ReplaceAndHighlight.xlsx";
	wstring outputFile = output_path + L"ReplaceAndHighlight_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	auto ranges = sheet->FindAllString(L"Total", true, true);
	//for (auto range : ranges)
	//{
	for (int i = 0; i < ranges->GetCount(); i++)
	{
		intrusive_ptr<CellRange> cr = ranges->GetItem(i);
		//reset the text, in other words, replace the text
		cr->SetText(L"Sum");
		//set the color
		cr->GetStyle()->SetColor(Spire::Xls::Color::GetYellow());
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

