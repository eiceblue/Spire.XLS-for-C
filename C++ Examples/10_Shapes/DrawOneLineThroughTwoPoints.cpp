#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawOneLineThroughTwoPoints.xlsx";

	//Create a workbook
	intrusive_ptr<Workbook> workbook = new Workbook();

	//Get the first worksheet
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	//1)Draw a line according to relative position
	intrusive_ptr<XlsLineShape> line1 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
	line1->SetLeftColumn(3);
	line1->SetTopRow(3);
	line1->SetLeftColumnOffset(0);
	line1->SetTopRowOffset(0);

	line1->SetRightColumn(4);
	line1->SetBottomRow(5);
	line1->SetRightColumnOffset(0);
	line1->SetBottomRowOffset(0);

	//2)Draw a line according to absolute position(pixels).
	intrusive_ptr<XlsLineShape> line2 = dynamic_pointer_cast<XlsLineShape>(sheet->GetTypedLines()->AddLine());
	intrusive_ptr<Point> startPoint = new Point();
	startPoint->SetX(30), startPoint->SetY(50);
	line2->SetStartPoint(startPoint);
	intrusive_ptr<Point> endPoint = new Point();
	endPoint->SetX(20), endPoint->SetY(80);
	line2->SetEndPoint(endPoint);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}
