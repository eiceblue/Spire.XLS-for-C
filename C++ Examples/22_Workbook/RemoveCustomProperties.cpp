#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"templateAz.xlsx";
	std::wstring outputFile = output_path + L"RemoveCustomProperties.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Retrieve a list of all custom document properties of the Excel file
	intrusive_ptr<ICustomDocumentProperties> customDocumentProperties = workbook->GetCustomDocumentProperties();

	//Remove "Editor" custom document property
	customDocumentProperties->Remove(L"Editor");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}