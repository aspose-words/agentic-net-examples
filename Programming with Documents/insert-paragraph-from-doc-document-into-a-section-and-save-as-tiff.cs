using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class InsertParagraphAndSaveAsTiff
{
    static void Main()
    {
        // Load the source DOC document that contains the paragraph to copy.
        Document srcDoc = new Document("Source.docx");

        // Create a new destination document.
        Document dstDoc = new Document();

        // Create a new section and add it to the destination document.
        Section newSection = new Section(dstDoc);
        dstDoc.AppendChild(newSection);

        // Import the first paragraph from the source document into the destination document.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(srcParagraph, true);

        // Append the imported paragraph to the new section's body.
        newSection.Body.AppendChild(importedParagraph);

        // Save the resulting document as a multi‑page TIFF image.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Optional: set resolution or other image options here if needed.
        dstDoc.Save("Result.tiff", tiffOptions);
    }
}
