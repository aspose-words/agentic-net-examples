using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToEpubConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the EPUB file will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.epub";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Save the document in EPUB format.
        doc.Save(outputPath, SaveFormat.Epub);
    }
}
