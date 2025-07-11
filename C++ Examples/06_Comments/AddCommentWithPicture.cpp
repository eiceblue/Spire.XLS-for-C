#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"Logo.png";
	wstring outputFile = outputFolder + L"AddCommentWithPicture_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	sheet->GetRange(L"C6")->SetText(L"E-iceblue");
	//Add the comment
	intrusive_ptr<ExcelComment> comment = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C6"))->AddExcelComment();
	//Load the image file
	intrusive_ptr<Stream> image = new Stream(inputFile.c_str());

	comment->GetFill()->CustomPicture(image, L"logo.png");
	//Set the height and width of comment
	comment->SetVisible(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
