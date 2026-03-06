using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank document that will receive the paragraph.
        Document dstDoc = new Document();

        // Ensure the destination document has at least one section, body, and paragraph.
        dstDoc.EnsureMinimum();

        // Retrieve the first non‑empty paragraph from the source document.
        // (Skip the automatically added empty last paragraph if present.)
        Paragraph srcParagraph = srcDoc.FirstSection.Body.Paragraphs[0];

        // Import the paragraph node into the destination document, preserving its formatting.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the first section in the destination document.
        dstDoc.FirstSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as an SVG file.
        dstDoc.Save("Result.svg", SaveFormat.Svg);
    }
}
