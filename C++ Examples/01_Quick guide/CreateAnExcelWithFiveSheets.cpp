#include "pch.h"

using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateAnExcelWithFiveSheets_result.xlsx";


	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->CreateEmptySheets(5);
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	for (int i = 0; i < 5; i++)
	{
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
		sheet->SetName((L"Sheet" + std::to_wstring(i)).c_str());
		for (int row = 1; row <= 150; row++)
		{
			for (int col = 1; col <= 50; col++)
			{
				dynamic_pointer_cast<CellRange>(sheet->GetRange(row, col))->SetText((L"row" + std::to_wstring(row) + L" col" + std::to_wstring(col)).c_str());
			}
		}
	}
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);


}

