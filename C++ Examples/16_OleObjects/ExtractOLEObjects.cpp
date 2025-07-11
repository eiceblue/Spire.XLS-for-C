#include "pch.h"
using namespace Spire::Xls;

// Convert a wide string to a standard string
std::string wstring2string(const std::wstring& wstr)
{
    std::string result;
    result.reserve(wstr.size());
    for (size_t i = 0; i < wstr.size(); ++i)
    {
        result += static_cast<char>(wstr[i] & 0xFF);
    }
    return result;
}

// Write all bytes to a file
void WriteAllBytes(std::wstring filePath, std::vector<byte> data)
{
    std::ofstream outFile(wstring2string(filePath), std::ios::out | std::ofstream::binary);
    outFile.write((char*)(&data[0]), data.size() * sizeof(byte));
    outFile.close();
}

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExtractOle2.xlsx";
	wstring outputDocFile = output_path + L"ExtractOLEObjects.docx";
	wstring outputPdfFile = output_path + L"ExtractOLEObjects.pdf";
	wstring outputPptFile = output_path + L"ExtractOLEObjects.pptx";


	// Create a Workbook object
	intrusive_ptr<Workbook> workbook = new Workbook();

	// Load the input file into the workbook
	workbook->LoadFromFile(inputFile.c_str());

	// Get the first worksheet from the workbook
	intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));

	// Check if the worksheet has OleObjects (embedded objects)
	if (sheet->GetHasOleObjects())
	{
		// Get the count of OleObjects in the worksheet
		int count = sheet->GetOleObjects()->GetCount();
		for (int i = 0; i < count; i++)
		{
			// Get the OleObject at the specified index
			auto Object = sheet->GetOleObjects()->GetItem(i);
			OleObjectType type = Object->GetObjectType();
			switch (type)
			{
				// Handle Word document objects
				case OleObjectType::WordDocument:
					WriteAllBytes(outputDocFile, Object->GetOleData());
					break;

				// Handle Adobe Acrobat document objects
				case OleObjectType::AdobeAcrobatDocument:
					WriteAllBytes(outputPdfFile, Object->GetOleData());
					break;

				// Handle PowerPoint slide objects
				case OleObjectType::PowerPointSlide:
					WriteAllBytes(outputPptFile, Object->GetOleData());
					break;

				default:
					break;
			}
		}
	}

	// Dispose the workbook to release resources
	workbook->Dispose();
}