#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"TrackChanges.xlsx";
	std::wstring outputFile = output_path + L"AcceptOrRejectTrackedChanges.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Accept the changes or reject the changes.
	//workbook.AcceptAllTrackedChanges();
	workbook->RejectAllTrackedChanges();

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

}