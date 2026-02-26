using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document srcDoc = new Document("Source.doc");

        // Retrieve the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new blank destination document.
        Document dstDoc = new Document();
        // Ensure the document has at least one section, body, and paragraph.
        dstDoc.EnsureMinimum();

        // Import the source paragraph into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Insert the imported paragraph after the last paragraph of the destination document.
        Paragraph lastParagraph = dstDoc.FirstSection.Body.LastParagraph;
        lastParagraph.ParentNode.InsertAfter(importedParagraph, lastParagraph);

        // Save the resulting document as plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        dstDoc.Save("Result.txt", txtOptions);
    }
}
