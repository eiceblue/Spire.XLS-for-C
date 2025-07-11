#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring outputFolder = OUTPUTPATH;
	wstring outputFile = outputFolder + L"WrapOrUnwrapTextInCells_out.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Wrap the excel text;
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->SetText(L"e-iceblue is in facebook and welcome to like us");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->GetStyle()->SetWrapText(true);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D1"))->SetText(L"e-iceblue is in twitter and welcome to follow us");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D1"))->GetStyle()->SetWrapText(true);

	//Unwrap the excel text;
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->SetText(L"http://www.facebook.com/pages/e-iceblue/139657096082266");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2"))->GetStyle()->SetWrapText(false);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2"))->SetText(L"https://twitter.com/eiceblue");
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"D2"))->GetStyle()->SetWrapText(false);

	//Set the text color of Range["C1:D1"]
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1:D1"))->GetStyle()->GetFont()->SetSize(15);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1:D1"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetBlue());
	//Set the text color of Range["C2:D2"]
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:D2"))->GetStyle()->GetFont()->SetSize(15);
	dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C2:D2"))->GetStyle()->GetFont()->SetColor(Spire::Xls::Color::GetDeepSkyBlue());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
	
}
