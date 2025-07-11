#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring outputFile = output + L"GetExcelPaperDimensions.txt";
	wfstream ofs;

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	wstring* content = new wstring();

	//Get the dimensions of A2 paper.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::A2Paper);
	content->append(L"A2Paper: " + to_wstring(dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup())->GetPageWidth()) + L" x " + to_wstring(dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup())->GetPageHeight()) + L"\r\n");

	//Get the dimensions of A3 paper.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperA3);
	content->append(L"PaperA3: " + to_wstring(dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup())->GetPageWidth()) + L" x " + to_wstring(dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup())->GetPageHeight()) + L"\r\n");

	//Get the dimensions of A4 paper.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperA4);
	content->append(L"PaperA4: " + to_wstring(dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup())->GetPageWidth()) + L" x " + to_wstring(dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup())->GetPageHeight()) + L"\r\n");

	//Get the dimensions of paper letter.
	sheet->GetPageSetup()->SetPaperSize(PaperSizeType::PaperLetter);
	content->append(L"PaperLetter: " + to_wstring(dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup())->GetPageWidth()) + L" x " + to_wstring(dynamic_pointer_cast<PageSetup>(sheet->GetPageSetup())->GetPageHeight()) + L"\r\n");

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}