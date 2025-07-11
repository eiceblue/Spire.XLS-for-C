#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"ReadStream.xlsx";
	std::wstring outputFile = output_path + L"ReadStream.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Open excel from a stream
	ifstream inputf(inputFile.c_str(), ios::in | ios::binary);
	intrusive_ptr<Stream> stream = new Stream(inputf);

	workbook->LoadFromStream(stream);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}