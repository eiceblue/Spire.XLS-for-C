#include "pch.h"
using namespace Spire::Xls;

// Convert a wide string (wstring) to a regular string (string)
string wstring2string(const std::wstring& wstr)
{
	std::string result;
	result.reserve(wstr.size());
	for (size_t i = 0; i < wstr.size(); ++i)
	{
		result += static_cast<char>(wstr[i] & 0xFF);
	}
	return result;
}

// Write content to a file specified by the outputFilePath
void WriteAllText(wstring outputFilePath, wstring contentStr)
{
	std::wofstream outputFile(wstring2string(outputFilePath), std::ios::out | std::ios::binary);

	outputFile.imbue(std::locale(""));

	outputFile << contentStr;

	outputFile.close();
}
	
int main() {
	wstring data_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = data_path + L"CellLinkedRangeLocal.xlsx";
	wstring outputFile = output_path + L"output.xlsx";

	// Create an intrusive pointer to a Workbook object
	intrusive_ptr<Workbook>workbook = new Workbook();

	// Create a string builder to store text
	wstring* sb = new wstring();

	// Load the workbook from the input file
	workbook->LoadFromFile(inputFile.c_str());

	// Get the first worksheet from the workbook
	intrusive_ptr<Worksheet>sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Get the collection of preset geometric shapes from the worksheet
	intrusive_ptr<PrstGeomShapeCollection>prstGeomShapeCollection = sheet->GetPrstGeomShapes();

	// Get the shape with the name "Yesterday" from the collection
	auto shape = prstGeomShapeCollection->Get(L"Yesterday");

	// Get the cell address linked to the shape
	std::wstring cellAddress = shape->GetLinkedCell()->GetRangeAddress();

	// Append the cell address to the string builder
	sb->append(cellAddress + L"\n");

	// Get the shape with the name "NewShapes" from the collection
	shape = prstGeomShapeCollection->Get(L"NewShapes");

	// Get the cell address linked to the shape
	cellAddress = shape->GetLinkedCell()->GetRangeAddress();

	// Append the cell address to the string builder
	sb->append(cellAddress);

	// Write the contents of the string builder to the output file
	WriteAllText(outputFile, *sb);

	// Dispose of the workbook object
	workbook->Dispose();
}

