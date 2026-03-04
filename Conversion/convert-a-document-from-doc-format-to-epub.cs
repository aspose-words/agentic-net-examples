using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputFile = @"C:\Docs\SourceDocument.doc";

        // Path where the EPUB file will be saved.
        string outputFile = @"C:\Docs\ConvertedDocument.epub";

        // Load the existing DOC document.
        Document doc = new Document(inputFile);

        // Save the document in EPUB format.
        doc.Save(outputFile, SaveFormat.Epub);
    }
}
