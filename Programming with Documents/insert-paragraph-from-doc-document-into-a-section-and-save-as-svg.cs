using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertParagraphAndSaveSvg
{
    static void Main()
    {
        // Load the source document that contains the paragraph to be inserted.
        Document srcDoc = new Document("Source.docx");

        // Retrieve the first paragraph from the source document.
        // Adjust the index if you need a different paragraph.
        Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

        // Create a new destination document.
        Document dstDoc = new Document();

        // Use DocumentBuilder to work with the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Ensure we are at the end of the first (and only) section.
        builder.MoveToDocumentEnd();

        // Import the source paragraph into the destination document.
        // ImportNode clones the node and resolves any references.
        Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

        // Insert the imported paragraph at the current cursor position.
        builder.InsertNode(importedParagraph);

        // Configure SVG save options (optional: render text as placed glyphs).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
        };

        // Save the resulting document as an SVG file.
        dstDoc.Save("Result.svg", svgOptions);
    }
}
