using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template.
        Document doc = new Document("Template.docx");

        // Evaluate all fields so that expression results are reflected in the document.
        doc.UpdateFields();

        // Configure SVG save options:
        // - No page border.
        // - Fit the SVG to the viewport (width/height = 100%).
        // - Render text as placed glyphs (curves) to ensure the SVG looks the same on any machine.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            ShowPageBorder = false,
            FitToViewPort = true,
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
        };

        // Save each page of the document as a separate SVG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            svgOptions.PageSet = new PageSet(pageIndex);
            string outputFile = $"Page_{pageIndex + 1}.svg";
            doc.Save(outputFile, svgOptions);
        }
    }
}
