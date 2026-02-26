using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.docx");

        // Create a new destination document.
        Document dstDoc = new Document();
        dstDoc.EnsureMinimum(); // Guarantees at least one section, body, and paragraph.

        // Retrieve the first non‑empty paragraph from the source document.
        Paragraph srcParagraph = (Paragraph)srcDoc.GetChild(NodeType.Paragraph, 0, true);
        if (srcParagraph.IsEndOfSection && !srcParagraph.HasChildNodes)
        {
            // Skip the automatically added empty end‑of‑section paragraph.
            srcParagraph = (Paragraph)srcDoc.GetChild(NodeType.Paragraph, 1, true);
        }

        // Import the paragraph into the destination document, preserving its formatting.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Insert the imported paragraph at the end of the first section's body.
        Section firstSection = dstDoc.FirstSection;
        firstSection.Body.InsertAfter(importedParagraph, firstSection.Body.LastParagraph);

        // Render the first page of the document to a PNG image.
        dstDoc.Save("Result.png", SaveFormat.Png);
    }
}
