#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"WorkbookToHTML.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"WorkbookToHTML.html";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Save to file.
		workbook->SaveToHtml(outputFile.c_str());
		workbook->Dispose();
}
