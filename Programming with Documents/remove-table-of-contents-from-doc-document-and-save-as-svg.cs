using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class RemoveTocAndSaveAsSvg
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting SVG will be saved.
        string outputPath = @"C:\Docs\ResultDocument.svg";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Remove all Table of Contents (TOC) fields from the document.
        // FieldType.FieldTOC identifies a TOC field.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // Remove the field and its result.
                field.Remove();
            }
        }

        // Optionally clean up any empty paragraphs left after removing TOCs.
        doc.Cleanup();

        // Configure SVG save options.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            // Fit the SVG to the viewport (browser window).
            FitToViewPort = true,

            // Do not draw a page border around the SVG.
            ShowPageBorder = false,

            // Render text as placed glyphs (curves) for better visual fidelity.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
        };

        // Save the modified document as an SVG file.
        doc.Save(outputPath, svgOptions);
    }
}
