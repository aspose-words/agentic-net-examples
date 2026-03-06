using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template.
        Document doc = new Document("Template.pdf");

        // Configure SVG save options.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            FitToViewPort = true,                     // Make SVG fill the viewport.
            ShowPageBorder = false,                   // No page border in the output.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs // Render text as curves.
        };

        // Export each page of the PDF as a separate SVG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            svgOptions.PageSet = new PageSet(pageIndex); // Select the current page.
            string outputPath = $"Page_{pageIndex + 1}.svg";
            doc.Save(outputPath, svgOptions);           // Save using the configured options.
        }
    }
}
