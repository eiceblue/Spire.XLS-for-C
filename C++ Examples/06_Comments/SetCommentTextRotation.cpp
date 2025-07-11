#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CellValues.xlsx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetCommentTextRotation.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Create Excel font
	intrusive_ptr<ExcelFont> font = workbook->CreateExcelFont();
	font->SetFontName(L"Arial");
	font->SetSize(11);
	font->SetKnownColor(ExcelColors::Orange);

	//Add the comment
	intrusive_ptr<CellRange> range = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"E1"));
	range->GetComment()->SetText(L"This is a comment");
	wstring text = range->GetComment()->GetText();
	range->GetComment()->GetRichText()->SetFont(0, (text.size() - 1), font);

	// Set its vertical and horizontal alignment 
	range->GetComment()->SetVAlignment(CommentVAlignType::Center);
	range->GetComment()->SetHAlignment(CommentHAlignType::Right);

	//Set the comment text rotation
	range->GetComment()->SetTextRotation(TextRotationType::LeftToRight);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
