#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CommentSample.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"RemoveComment.xlsx";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get all comments of the first sheet
		intrusive_ptr<XlsCommentsCollection> comments = dynamic_pointer_cast<XlsCommentsCollection>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetComments());
		//Change the content of the first comment
		comments->Get(0)->SetText(L"This comment has been changed.");
		//Remove the second comment
		comments->Get(1)->Remove();

		//Save to file.
		workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
		workbook->Dispose();
}
