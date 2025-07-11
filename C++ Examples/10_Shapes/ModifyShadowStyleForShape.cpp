#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Xls_5.xlsx";
	wstring outputFile = output_path + L"ModifyShadowStyleForShape.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Get the third shape from the worksheet.
	intrusive_ptr<IPrstGeomShape> shape = sheet->GetPrstGeomShapes()->Get(2);

	//Set the shadow style for the shape.
	shape->GetShadow()->SetAngle(90);
	shape->GetShadow()->SetTransparency(30);
	shape->GetShadow()->SetDistance(10);
	shape->GetShadow()->SetSize(130);
	shape->GetShadow()->SetColor(Spire::Xls::Color::GetYellow());
	shape->GetShadow()->SetBlur(30);
	shape->GetShadow()->SetHasCustomStyle(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
