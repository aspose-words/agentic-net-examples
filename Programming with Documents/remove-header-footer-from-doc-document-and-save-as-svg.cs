using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeaderFooterAndSaveSvg
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting SVG will be saved.
        string outputPath = @"C:\Docs\ResultDocument.svg";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.HeadersFooters.Clear();
        }

        // Configure SVG save options (optional customizations).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            // Do not draw a page border around the SVG.
            ShowPageBorder = false,
            // Render text as placed glyphs (curves) to avoid selectable text.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            // Fit the SVG to the viewport for easier display.
            FitToViewPort = true
        };

        // Save the document as SVG using the configured options.
        doc.Save(outputPath, svgOptions);
    }
}
