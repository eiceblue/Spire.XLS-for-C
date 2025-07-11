#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	wstring outputFile = output_path + L"AddLineShape.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0)));

	//Add shape line1
	intrusive_ptr<ILineShape> line1 = sheet->GetLines()->AddLine(10, 2, 200, 1, LineShapeType::Line);
	//Set dash style type
	line1->SetDashStyle(ShapeDashLineStyleType::Solid);
	//Set color
	line1->SetColor(Spire::Xls::Color::GetCadetBlue());
	//Set weight
	line1->SetWeight(2.0f);
	//Set end arrow style type
	line1->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);

	//Add shape line2
	intrusive_ptr<ILineShape> line2 = sheet->GetLines()->AddLine(12, 2, 200, 1, LineShapeType::CurveLine);
	line2->SetDashStyle(ShapeDashLineStyleType::Dotted);
	line2->SetColor(Spire::Xls::Color::GetOrangeRed());
	line2->SetWeight(2.0f);

	//Add shape line3
	intrusive_ptr<ILineShape> line3 = sheet->GetLines()->AddLine(14, 2, 200, 1, LineShapeType::ElbowLine);
	line3->SetDashStyle(ShapeDashLineStyleType::DashDotDot);
	line3->SetColor(Spire::Xls::Color::GetPurple());
	line3->SetWeight(2.0f);

	//Add shape line4
	intrusive_ptr<ILineShape> line4 = sheet->GetLines()->AddLine(16, 2, 200, 1, LineShapeType::LineInv);
	line4->SetDashStyle(ShapeDashLineStyleType::Dashed);
	line4->SetColor(Spire::Xls::Color::GetGreen());
	line4->SetWeight(2.0f);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
