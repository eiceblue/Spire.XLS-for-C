#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConversionSample1.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToSVG";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		for (int i = 0; i < workbook->GetWorksheets()->GetCount(); i++)
		{
			intrusive_ptr<Stream> fileStream = new Stream();
			dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(i))->ToSVGStream(fileStream, 0, 0, 0, 0);
			fileStream->Save((outputFile + L"sheet-" + std::to_wstring(i) + L".svg").c_str());
			//fileStream->~Stream();
		}

		workbook->Dispose();
}
