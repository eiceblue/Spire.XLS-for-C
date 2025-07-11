#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ShapeToImage.xlsx";
	wstring outputFile = output_path + L"ShapeToImage.png";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the first shape from the first worksheet
	intrusive_ptr<XlsShape> shape = dynamic_pointer_cast<XlsShape>(sheet->GetPrstGeomShapes()->Get(0));

	//Save the shape to a image
	intrusive_ptr<Stream> img = shape->SaveToImage();
	img->Save(outputFile.c_str());
	img->Close();

	workbook->Dispose();
}
