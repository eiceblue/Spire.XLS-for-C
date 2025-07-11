#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"GetDefaultRowAndColumnCount_out.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Clear all worksheets
	workbook->GetWorksheets()->Clear();

	//Create a new worksheet
	intrusive_ptr<Worksheet> sheet = workbook->CreateEmptySheet();
	std::wstring* builder = new std::wstring();
	//Get row and column count
	int rowCount = sheet->GetRows()->GetCount();
	int columnCount = sheet->GetColumns()->GetCount();

	builder->append(L"The default row count is :" + std::to_wstring(rowCount));
	builder->append(L"The default column count is :" + std::to_wstring(columnCount));

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << builder << endl;
	ofs.close();
	workbook->Dispose();
}
