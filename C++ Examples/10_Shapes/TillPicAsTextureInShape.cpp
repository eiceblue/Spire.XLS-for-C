#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"TillPicAsTextureInShape.xlsx";
	wstring outputFile = output_path + L"TillPicAsTextureInShape.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the first shape
	intrusive_ptr<IPrstGeomShape> shape = sheet->GetPrstGeomShapes()->Get(0);

	//Fill shape with texture
	shape->GetFill()->SetFillType(ShapeFillType::Texture);

	//Custom texture with picture
	shape->GetFill()->CustomTexture((input_path  + L"Logo.png").c_str());

	//Tile pciture as texture 
	shape->GetFill()->SetTile(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
