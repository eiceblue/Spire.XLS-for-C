#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ReadImages.xlsx";
   	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToHtmlStream.html";

		//Create a workbook
		intrusive_ptr<Workbook> workbook = new Workbook();

		//Load the Excel document from disk
		workbook->LoadFromFile(inputFile.c_str());

		//Get the first worksheet
		intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

		//Set the html options
		intrusive_ptr<HTMLOptions> options = new HTMLOptions();
		options->SetImageEmbedded(true);
		//Save sheet to html stream
		intrusive_ptr<Stream> stream = new Stream();

		//Save to file.
		sheet->SaveToHtml(stream, options);
		workbook->Dispose();

		stream->Save(outputFile.c_str());
}
