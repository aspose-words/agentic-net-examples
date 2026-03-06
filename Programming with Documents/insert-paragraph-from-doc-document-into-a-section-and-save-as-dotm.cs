using System;
using Aspose.Words;

class InsertParagraphIntoSection
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank document that will receive the paragraph.
        Document dstDoc = new Document();
        dstDoc.EnsureMinimum(); // Guarantees at least one section and body.

        // Retrieve the first (or any) paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document, preserving its formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the destination document's first section.
        dstDoc.FirstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as a DOTM (Word macro‑enabled template).
        dstDoc.Save("Result.dotm");
    }
}
