#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateAnExcelWithOneSheet_result.xlsx";


	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();
	workbook->CreateEmptySheets(1);
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
	for (int row = 1; row <= 10000; row++)
	{
		for (int col = 1; col <= 30; col++)
		{
			dynamic_pointer_cast<CellRange>(sheet->GetRange(row, col))->SetText((L"row" + std::to_wstring(row) + L" col" + std::to_wstring(col)).c_str());
		}
	}
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2010);


}

