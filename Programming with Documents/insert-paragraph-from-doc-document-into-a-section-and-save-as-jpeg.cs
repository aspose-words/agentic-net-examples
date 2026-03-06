using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC file that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.doc");

        // Create a new blank target document.
        Document targetDoc = new Document();

        // Ensure the target document has at least one section, body and paragraph.
        targetDoc.EnsureMinimum();

        // Retrieve the first non‑empty paragraph from the source document.
        Paragraph srcParagraph = (Paragraph)srcDoc.GetChild(NodeType.Paragraph, 0, true);

        // Import the paragraph into the target document, preserving its original formatting.
        NodeImporter importer = new NodeImporter(srcDoc, targetDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the body of the first section in the target document.
        targetDoc.FirstSection.Body.AppendChild(importedParagraph);

        // Configure image save options for JPEG output.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Optional: set JPEG quality (0‑100). Higher value = better quality, larger file.
            JpegQuality = 90
        };

        // Save the target document as a JPEG image (only the first page is rendered).
        targetDoc.Save("Result.jpg", jpegOptions);
    }
}
