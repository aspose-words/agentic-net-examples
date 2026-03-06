using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains expression fields (e.g., MERGEFIELD, IF, etc.).
        Document doc = new Document(@"C:\Input\Template.docx");

        // Ensure that all fields are evaluated so the resulting SVG contains the computed values.
        doc.UpdateFields();

        // Configure SVG save options.
        // - FitToViewPort makes the SVG fill the container (width/height = 100%).
        // - ShowPageBorder = false removes the page outline.
        // - TextOutputMode = UsePlacedGlyphs renders text as curves (no selectable text, but preserves appearance).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            FitToViewPort = true,
            ShowPageBorder = false,
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
        };

        // Save the document as an SVG file using the configured options.
        doc.Save(@"C:\Output\Result.svg", svgOptions);
    }
}
