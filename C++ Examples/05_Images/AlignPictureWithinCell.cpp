#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring inputFolder = DATAPATH;
	wstring outputFolder = OUTPUTPATH;
	wstring inputFile = inputFolder + L"intro_xls.png";
	wstring outputFile = outputFolder + L"AlignPictureWithinCell_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	sheet->GetRange(L"A1")->SetText(L"Align Picture Within A Cell:");
	sheet->GetRange(L"A1")->GetStyle()->SetVerticalAlignment(VerticalAlignType::Top);
	intrusive_ptr<ExcelPicture> picture = ExcelPicture::Dynamic_cast<ExcelPicture>(sheet->GetPictures()->Add(1, 1, inputFile.c_str()));

	//Adjust the column width and row height so that the cell can contain the picture.
	sheet->GetColumns()->GetItem(0)->SetColumnWidth(40);
	sheet->GetRows()->GetItem(0)->SetRowHeight(200);

	//Vertically and horizontally align the image.
	picture->SetLeftColumnOffset(100);
	picture->SetTopRowOffset(25);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
