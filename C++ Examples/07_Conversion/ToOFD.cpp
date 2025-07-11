#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"ToOFD.xlsx";
	wstring outputFile = output_path + L"ToOFD.ofd";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Save the workbook to the specified output file with the specified file format (FileFormat::OFD).
	workbook->SaveToFile(outputFile.c_str(), FileFormat::OFD);

	// Dispose of the workbook object.
	workbook->Dispose();
}

