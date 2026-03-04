using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source file (any supported format, e.g., PDF, DOCX, etc.)
        string sourcePath = @"C:\Input\sample.pdf";

        // Load the document using the constructor that accepts a file name.
        // The format is detected automatically.
        Document doc = new Document(sourcePath);

        // Path where the converted DOC file will be saved.
        string outputPath = @"C:\Output\sample.doc";

        // Save the loaded document in the legacy DOC format.
        doc.Save(outputPath, SaveFormat.Doc);
    }
}
