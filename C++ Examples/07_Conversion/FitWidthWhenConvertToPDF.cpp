#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.xlsx";
   	 wstring output_path = OUTPUTPATH;
   	 wstring outputFile = output_path + L"FitWidthWhenConvertToPDF.pdf";

	 //Create a workbook
	 intrusive_ptr<Workbook> workbook = new Workbook();

	 //Load the Excel document from disk
	 workbook->LoadFromFile(inputFile.c_str());

	 for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
	 {
		 intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i));
		 //Auto fit page height
		 sheet->GetPageSetup()->SetFitToPagesTall(0);
		 //Fit one page width
		 sheet->GetPageSetup()->SetFitToPagesWide(1);
	 }

	 //Save to file.
	 workbook->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	 workbook->Dispose();
}
