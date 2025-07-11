#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SetShapeOrder.xlsx";
	wstring outputFile = output_path + L"SetShapeOrder.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Bring the picture forward one level
	dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0))->GetPictures()->Get(0)->ChangeLayer(ShapeLayerChangeType::BringForward);

	//Bring the image in fron of all other objects
	dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(1))->GetPictures()->Get(0)->ChangeLayer(ShapeLayerChangeType::BringToFront);

	//Send the shape back one level
	intrusive_ptr<XlsShape> shape = dynamic_pointer_cast<XlsShape>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(2))->GetPrstGeomShapes()->Get(1));
	shape->ChangeLayer(ShapeLayerChangeType::SendBackward);

	//Send the shape behind all other objects
	shape = dynamic_pointer_cast<XlsShape>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(3))->GetPrstGeomShapes()->Get(1));
	shape->ChangeLayer(ShapeLayerChangeType::SendToBack);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
