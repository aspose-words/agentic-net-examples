using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source document (any supported format).
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the document using the Document(string) constructor.
        Document doc = new Document(sourcePath);

        // Save the loaded document as DOCX.
        string outputPath = @"C:\Docs\ConvertedDocument.docx";
        doc.Save(outputPath);
    }
}
