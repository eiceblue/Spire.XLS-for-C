#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"GetFreezePaneRange.xlsx";
	std::wstring outputFile = output_path + L"GetFreezePaneRange.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//The row and column index of the frozen pane is passed through the out parameter. 
	//If it returns to 0, it means that it is not frozen
	int rowIndex = sheet->GetFreezePanes()[0];
	int colIndex = sheet->GetFreezePanes()[1];

	wstring range = L"Row index: " + to_wstring(rowIndex) + L", column index: " + to_wstring(colIndex);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << range << endl;
	ofs.close();
	workbook->Dispose();
}