using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingToPng
{
    // Simple data source for the LINQ Reporting Engine.
    public class ReportData
    {
        public string Title { get; set; }
        public int Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the PDF template that contains LINQ Reporting expressions.
            // The template file must exist at the specified path.
            Document template = new Document(@"C:\Templates\ReportTemplate.pdf");

            // Prepare the data source.
            var data = new ReportData
            {
                Title = "Quarterly Sales",
                Value = 12500
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // Allow missing members to avoid runtime errors if the template references non‑existent fields.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.BuildReport(template, data, "ds"); // "ds" is the name used in the template.

            // Configure image save options to render the first page as PNG.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0),
                // Optional: set resolution (dpi) for higher quality.
                Resolution = 300
            };

            // Save the rendered page as a PNG image.
            string outputPath = @"C:\Output\ReportPage1.png";
            template.Save(outputPath, pngOptions);
        }
    }
}
