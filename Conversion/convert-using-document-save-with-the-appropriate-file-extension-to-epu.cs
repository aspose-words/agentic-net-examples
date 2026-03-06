using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path to the output EPUB file. The .epub extension determines the save format.
        string outputPath = @"C:\Docs\SourceDocument.epub";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Save the document as EPUB. The format is inferred from the file extension.
        doc.Save(outputPath);
    }
}
