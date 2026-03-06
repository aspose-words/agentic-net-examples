using System;
using Aspose.Words;

class InsertParagraphIntoSection
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank destination document.
        Document dstDoc = new Document();
        dstDoc.EnsureMinimum(); // Guarantees at least one section, body and paragraph.

        // Get the paragraph you want to insert (first paragraph of the first section as an example).
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph node from the source document into the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Insert the imported paragraph into the desired section of the destination document.
        // Here we insert it after the last paragraph of the first section's body.
        Section targetSection = dstDoc.FirstSection;
        targetSection.Body.InsertAfter(importedParagraph, targetSection.Body.LastParagraph);

        // Save the resulting document as a Word template (.dot).
        dstDoc.Save("Result.dot"); // Extension .dot automatically selects SaveFormat.Dot.
    }
}
