#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToPostScript.xlsx";
    	wstring output_path = OUTPUTPATH;
   	wstring outputFile = output_path + L"ToPostScript.ps";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), FileFormat::PostScript);
	workbook->Dispose();
}
