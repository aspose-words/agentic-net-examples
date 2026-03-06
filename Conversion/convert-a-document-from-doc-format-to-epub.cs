using System;
using Aspose.Words;

class DocToEpubConverter
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputFile = "Document.doc";

        // Path where the EPUB file will be saved.
        string outputFile = "Document.epub";

        // Load the DOC document.
        Document doc = new Document(inputFile);

        // Save the document in EPUB format.
        doc.Save(outputFile, SaveFormat.Epub);
    }
}
