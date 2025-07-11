#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring inputFile = fn + L"ReadImages.xlsx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetCellDisplayedText_result.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Set value for B8
	intrusive_ptr<CellRange> cell = dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B8"));
	cell->SetNumberValue(0.012345);

	//Set the cell style
	intrusive_ptr<CellStyle> style = dynamic_pointer_cast<CellStyle>(cell->GetStyle());
	style->SetNumberFormat(L"0.00");

	//Get the cell value
	wstring cellValue = cell->GetValue();

	//Get the displayed text of the cell
	wstring displayedText = cell->GetDisplayedText();

	//Create StringBuilder to save 
	std::wstring* content = new std::wstring();

	//Set string format for displaying
	wstring result = L"B8 Value: " + cellValue + L"\r\nB8 displayed text: " + displayedText;

	//Add result string to StringBuilder
	content->append(result);

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << content << endl;
	ofs.close();
	workbook->Dispose();
}

