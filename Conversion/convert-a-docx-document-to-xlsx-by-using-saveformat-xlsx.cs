using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "Input.docx";

        // Path where the resulting XLSX file will be saved.
        string outputPath = "Output.xlsx";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Convert and save the document as an XLSX spreadsheet.
        doc.Save(outputPath, SaveFormat.Xlsx);
    }
}
