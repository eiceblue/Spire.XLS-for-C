#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"MacroSample.xls";
	std::wstring outputFile = output_path + L"DetectVBAMacros.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Detect if the Excel file contains VBA macros
	wstring value = L"";
	bool hasMacros = workbook->GetHasMacros();
	if (hasMacros)
	{
		value = L"Yes";
	}

	else
	{
		value = L"No";
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << value << endl;
	ofs.close();
	workbook->Dispose();
}