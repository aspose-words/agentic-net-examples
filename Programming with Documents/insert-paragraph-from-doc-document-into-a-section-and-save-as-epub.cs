using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveEpub
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be inserted.
        Document srcDoc = new Document("Source.doc");

        // Create a new (blank) destination document.
        Document dstDoc = new Document();

        // Obtain the first (or any) paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Prepare a NodeImporter to copy nodes from the source to the destination document,
        // preserving the source formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Import the paragraph node into the destination document.
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Insert the imported paragraph at the end of the first section of the destination document.
        Section targetSection = dstDoc.Sections[0];
        Paragraph lastParagraph = targetSection.Body.LastParagraph;
        targetSection.Body.InsertAfter(importedParagraph, lastParagraph);

        // Configure save options for EPUB output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,
            // Optional: set encoding, split criteria, etc., as needed.
        };

        // Save the resulting document as an EPUB file.
        dstDoc.Save("Result.epub", saveOptions);
    }
}
