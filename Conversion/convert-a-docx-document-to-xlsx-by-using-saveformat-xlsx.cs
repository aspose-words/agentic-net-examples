using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path to the destination XLSX file.
        string outputPath = "output.xlsx";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Save the document in XLSX format using SaveFormat.Xlsx.
        doc.Save(outputPath, SaveFormat.Xlsx);
    }
}
