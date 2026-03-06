using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Reporting;

namespace AsposeWordsSvgExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOC template that contains Aspose.Words reporting tags.
            Document doc = new Document("Template.docx");

            // Create a simple data source with properties referenced in the template.
            var dataSource = new
            {
                Title = "Sales Report",
                Total = 12500.75,
                Date = DateTime.Today
            };

            // Populate the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource);

            // Configure SVG save options.
            SvgSaveOptions svgOptions = new SvgSaveOptions
            {
                // Make the SVG fill the viewport (optional).
                FitToViewPort = true,
                // Remove the default page border.
                ShowPageBorder = false,
                // Render text as placed glyphs so the SVG does not depend on external fonts.
                TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
            };

            // Save the populated document as an SVG file.
            doc.Save("Result.svg", svgOptions);
        }
    }
}
