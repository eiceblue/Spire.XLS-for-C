#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"SampleB_2.xlsx";
	wstring outputFile = output_path + L"FindCellsWithStyleName_result.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the cell style name
	wstring styleName = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->GetCellStyleName();

	intrusive_ptr<CellRange> ranges = dynamic_pointer_cast<CellRange>(sheet->GetAllocatedRange());
	//for (auto cc : ranges->GetCells())
	//{
	for (int i = 0; i < ranges->GetCells()->GetCount(); i++)
	{
		intrusive_ptr<CellRange> cr = ranges->GetCells()->GetItem(i);
		//Find the cells which have the same style name
		if (!wcscmp(cr->GetCellStyleName(), styleName.c_str()))
		{
			//Set value
			cr->SetValue(L"Same style");
		}
	}

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

}

