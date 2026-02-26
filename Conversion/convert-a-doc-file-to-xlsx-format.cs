using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocToXlsx
{
    static void Main()
    {
        // Input DOC file path
        string inputPath = "input.doc";

        // Output XLSX file path
        string outputPath = "output.xlsx";

        // Load the DOC document
        Document doc = new Document(inputPath);

        // Save the document in XLSX format
        doc.Save(outputPath, SaveFormat.Xlsx);
    }
}
