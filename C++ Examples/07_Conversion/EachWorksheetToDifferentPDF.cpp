#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"EachWorksheetToDifferentPDFSample.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"EachWorksheetToPDF";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
		{
			intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
			wstring FileName = outputFile + sheet->GetName() + L".pdf";
			//Save the sheet to PDF
			sheet->SaveToPdf(FileName.c_str());
		}

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();
}
