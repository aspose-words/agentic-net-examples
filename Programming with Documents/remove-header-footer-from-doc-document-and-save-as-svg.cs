using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeaderFooterAndSaveSvg
{
    static void Main()
    {
        // Path to the source DOC/DOCX file.
        string inputFile = "input.docx";

        // Path where the resulting SVG will be saved.
        string outputFile = "output.svg";

        // Load the document from disk.
        Document doc = new Document(inputFile);

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.HeadersFooters.Clear();
        }

        // Configure SVG save options.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            // Do not draw a page border around each SVG page.
            ShowPageBorder = false,

            // Make the SVG fill the viewport (optional, improves display in browsers).
            FitToViewPort = true,

            // Render text as placed glyphs so that the SVG does not depend on external fonts.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
        };

        // Save the modified document as SVG using the configured options.
        doc.Save(outputFile, svgOptions);
    }
}
