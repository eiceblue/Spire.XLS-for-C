#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"CheckDialogSheetInXlsFile.xlsx";
	std::wstring outputFile = output_path + L"CheckDialogSheetInXlsFile.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	wstring* content = new wstring();

	//Check if the worksheet is a dialog sheet.
	if (sheet->GetType() == ExcelSheetType::DialogSheet)
	{
		content->append(L"Worksheet is a Dialog Sheet!");
	}
	else
	{
		content->append(L"Worksheet is not a Dialog Sheet!");
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}