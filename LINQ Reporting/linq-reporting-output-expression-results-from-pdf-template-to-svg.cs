using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportingToSvg
{
    // Simple data source class with properties used in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public decimal Amount { get; set; }
        public DateTime Date { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains LINQ Reporting tags, e.g. <<[data.Title]>>.
            const string pdfTemplatePath = @"C:\Templates\ReportTemplate.pdf";

            // Load the PDF template into a Document object.
            Document doc = new Document(pdfTemplatePath);

            // Prepare the data source.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Amount = 12345.67m,
                Date = DateTime.Today
            };

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the data source object, the third is the name used in the template.
            engine.BuildReport(doc, data, "data");

            // Configure SVG save options.
            SvgSaveOptions svgOptions = new SvgSaveOptions
            {
                // Render text as placed glyphs so the SVG contains vector outlines of the text.
                TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
                // Optional: remove page borders for a cleaner SVG.
                ShowPageBorder = false,
                // Optional: fit the SVG to the viewport.
                FitToViewPort = true
            };

            // Save the populated document as an SVG file.
            const string outputSvgPath = @"C:\Output\ReportResult.svg";
            doc.Save(outputSvgPath, svgOptions);
        }
    }
}
