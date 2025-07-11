#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddArrowLineToExcelFile.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//Add a Double Arrow and fill the line with solid color.
	auto line = sheet->GetTypedLines()->AddLine();
	line->SetTop(10);
	line->SetLeft(20);
	line->SetWidth(100);
	line->SetHeight(0);
	line->SetColor(Spire::Xls::Color::GetBlue());
	line->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrow);
	line->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);
	//Add an Arrow and fill the line with solid color.
	auto line_1 = sheet->GetTypedLines()->AddLine();
	line_1->SetTop(50);
	line_1->SetLeft(30);
	line_1->SetWidth(100);
	line_1->SetHeight(100);
	line_1->SetColor(Spire::Xls::Color::GetRed());
	line_1->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineNoArrow);
	line_1->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);

	//Add an Elbow Arrow Connector.
	intrusive_ptr<XlsLineShape> line3 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
	line3->SetLineShapeType(LineShapeType::ElbowLine);
	line3->SetWidth(30);
	line3->SetHeight(50);
	line3->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);
	line3->SetTop(100);
	line3->SetLeft(50);

	//Add an Elbow Double-Arrow Connector.
	intrusive_ptr<XlsLineShape> line2 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
	line2->SetLineShapeType(LineShapeType::ElbowLine);
	line2->SetWidth(50);
	line2->SetHeight(50);
	line2->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrow);
	line2->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrow);
	line2->SetLeft(120);
	line2->SetTop(100);

	//Add a Curved Arrow Connector.
	line3 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
	line3->SetLineShapeType(LineShapeType::CurveLine);
	line3->SetWidth(30);
	line3->SetHeight(50);
	line3->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrowOpen);
	line3->SetTop(100);
	line3->SetLeft(200);

	//Add a Curved Double-Arrow Connector.
	line2 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
	line2->SetLineShapeType(LineShapeType::CurveLine);
	line2->SetWidth(30);
	line2->SetHeight(50);
	line2->SetEndArrowHeadStyle(ShapeArrowStyleType::LineArrowOpen);
	line2->SetBeginArrowHeadStyle(ShapeArrowStyleType::LineArrowOpen);
	line2->SetLeft(250);
	line2->SetTop(100);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
