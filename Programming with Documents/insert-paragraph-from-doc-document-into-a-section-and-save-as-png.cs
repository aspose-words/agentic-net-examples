using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsPng
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Ensure the destination document has at least one section.
        // (A new Document already contains a default section, body, and paragraph.)
        Section targetSection = dstDoc.FirstSection;

        // Get the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph node into the destination document, preserving its formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the target section's body.
        targetSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as a PNG image (renders the first page).
        dstDoc.Save("Result.png", SaveFormat.Png);
    }
}
