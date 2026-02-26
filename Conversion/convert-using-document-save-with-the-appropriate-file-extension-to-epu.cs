using System;
using Aspose.Words;

class DocmToEpubConverter
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path to the output EPUB file. The .epub extension tells Aspose.Words to save in EPUB format.
        string outputPath = @"C:\Docs\ConvertedDocument.epub";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Save the document as EPUB. The format is inferred from the file extension.
        doc.Save(outputPath);
    }
}
