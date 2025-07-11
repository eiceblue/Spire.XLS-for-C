#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveDataValidation.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"RemoveDataValidation_out.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Create an array of rectangles, which is used to locate the ranges in worksheet.
		std::vector<intrusive_ptr<Spire::Xls::Rectangle>> rectangles(1);

		//Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
		rectangles[0] = Spire::Xls::Rectangle::FromLTRB(0, 0, 1, 2);

		//Remove validations in the ranges represented by rectangles.
		dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetDVTable()->Remove(rectangles);

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();
}
