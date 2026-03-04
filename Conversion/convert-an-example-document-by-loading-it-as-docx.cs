using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source document (can be any supported format, e.g., PDF, RTF, etc.).
        string sourcePath = @"C:\Docs\Example.pdf";

        // Path where the converted DOCX will be saved.
        string outputPath = @"C:\Docs\ExampleConverted.docx";

        // Load the source document. The Document constructor automatically detects the format.
        Document doc = new Document(sourcePath);

        // Save the loaded document as DOCX. The file extension determines the save format.
        doc.Save(outputPath);
    }
}
