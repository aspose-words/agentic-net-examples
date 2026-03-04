using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path to the target DOCX file.
        string outputPath = @"C:\Docs\ConvertedDocument.docx";

        // Load the existing DOC document.
        Document doc = new Document(inputPath);

        // Save the document in DOCX format.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
