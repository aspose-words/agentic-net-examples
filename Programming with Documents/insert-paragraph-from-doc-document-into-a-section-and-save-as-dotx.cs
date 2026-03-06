using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank destination document.
        Document dstDoc = new Document();
        dstDoc.EnsureMinimum(); // Guarantees at least one section, body and paragraph.

        // Get the first section of the destination where we will insert the paragraph.
        Section dstSection = dstDoc.FirstSection;

        // Select the paragraph to copy from the source document (first paragraph of the first section).
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document, preserving its formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the end of the destination section's body.
        dstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as a DOTX template.
        dstDoc.Save("Result.dotx", SaveFormat.Dotx);
    }
}
