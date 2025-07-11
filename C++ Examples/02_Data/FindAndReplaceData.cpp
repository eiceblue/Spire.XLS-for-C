#include"../pch.h"

using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"CreateTable.xlsx";
	wstring outputFile = output_path + L"FindAndReplaceData_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Find the L"Brazil" string
	auto ranges = sheet->FindAllString(L"Area", false, false);

	//Traverse the found ranges
	//for (auto range : ranges)
	//{
	for (int i = 0; i < ranges->GetCount(); i++)
	{
		intrusive_ptr<CellRange> cr = ranges->GetItem(i);
		//Replace it with L"China"
		cr->SetText(L"Area Code");
		//Highlight the color
		cr->GetStyle()->SetColor(Spire::Xls::Color::GetYellow());
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

}

