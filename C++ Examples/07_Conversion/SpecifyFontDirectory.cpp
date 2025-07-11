#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"ToPDFSample.xlsx";
	wstring inputFontFile = data_path + L"\\font\\";
	wstring outputFile = output_path + L"SpecifyFontDirectory.pdf";

	// Create a new Workbook object using intrusive_ptr smart pointer.
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the workbook from the specified input file.
	workbook->LoadFromFile(inputFile.c_str());

	// Create a vector to store custom font file paths as LPCWSTR_S (wide string) elements.
	vector<LPCWSTR_S> fonts;

	// Add the inputFontFile path to the fonts vector.
	fonts.push_back(inputFontFile.c_str());

	// Set the custom font file directory for the workbook using the fonts vector.
	workbook->SetCustomFontFileDirectory(fonts);

	// Save the workbook to the specified output file in PDF format.
	workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);

	// Dispose of the workbook object.
	workbook->Dispose();
}

