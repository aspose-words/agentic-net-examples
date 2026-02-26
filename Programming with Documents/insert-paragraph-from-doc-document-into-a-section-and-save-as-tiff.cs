using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveAsTiff
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be copied.
        Document srcDoc = new Document("Source.docx");

        // Create a new (empty) destination document.
        Document dstDoc = new Document();

        // Create a new section and add it to the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.AppendChild(newSection);

        // Ensure the new section has a body (it is created automatically, but we reference it for clarity).
        Body body = newSection.Body;

        // Get the first paragraph from the source document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Import the paragraph into the destination document, preserving its formatting.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the new section's body.
        body.AppendChild(importedParagraph);

        // Save the resulting document as a TIFF image (one page per image).
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        dstDoc.Save("Result.tiff", tiffOptions);
    }
}
