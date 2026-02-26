using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsJpeg
{
    static void Main()
    {
        // Path to the source DOC document that contains the paragraph to copy.
        string sourcePath = @"C:\Input\SourceDocument.doc";

        // Path where the resulting JPEG image will be saved.
        string outputPath = @"C:\Output\ResultImage.jpg";

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Create a new blank target document.
        Document targetDoc = new Document();

        // Get the first paragraph from the source document.
        Paragraph sourceParagraph = sourceDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph node into the target document, preserving its formatting.
        Node importedParagraph = targetDoc.ImportNode(sourceParagraph, true);

        // Append the imported paragraph to the first (and only) section of the target document.
        targetDoc.FirstSection.Body.AppendChild(importedParagraph);

        // Prepare image save options for JPEG format.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Optional: set JPEG quality (0-100). Higher value = better quality.
            JpegQuality = 90,
            // Render only the first page (the document has a single page with the paragraph).
            PageSet = new PageSet(0)
        };

        // Save the target document as a JPEG image.
        targetDoc.Save(outputPath, jpegOptions);
    }
}
