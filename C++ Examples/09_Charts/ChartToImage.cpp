#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"ChartToImage.xlsx";
	wstring outputFile = output_path + L"ChartToImage.png";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	intrusive_ptr<Stream> image = workbook->SaveChartAsImage(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0)), 0);
	image->Save(outputFile.c_str());
	image->Close();

	//Save to file.
	workbook->Dispose();
}
