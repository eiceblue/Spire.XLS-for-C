#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetShadowStyleForShape.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add an ellipse shape.
	intrusive_ptr<IPrstGeomShape> ellipse = sheet->GetPrstGeomShapes()->AddPrstGeomShape(5, 5, 150, 100, PrstGeomShapeType::Ellipse);

	//Set the shadow style for the ellipse.
	ellipse->GetShadow()->SetAngle(90);
	ellipse->GetShadow()->SetDistance(10);
	ellipse->GetShadow()->SetSize(150);
	ellipse->GetShadow()->SetColor(Spire::Xls::Color::GetGray());
	ellipse->GetShadow()->SetBlur(30);
	ellipse->GetShadow()->SetTransparency(1);
	ellipse->GetShadow()->SetHasCustomStyle(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
