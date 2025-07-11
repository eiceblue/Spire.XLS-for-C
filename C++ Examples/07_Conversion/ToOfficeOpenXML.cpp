#include "pch.h"
using namespace Spire::Xls;

int main() {
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToOfficeOpenXML.xml";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"A1"))->SetText(L"Hello World");
		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"B1"))->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);
		dynamic_pointer_cast<CellRange>(sheet->GetRange(L"C1"))->GetStyle()->SetKnownColor(ExcelColors::Gold);

		//Save to file.
		workbook->SaveAsXml(outputFile.c_str());
		workbook->Dispose();
}
