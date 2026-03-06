using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains Aspose.Words expressions (e.g., <<[Data.Property]>>)
            string templatePath = @"C:\Docs\Template.docx";

            // Path where the resulting fixed HTML will be saved
            string outputPath = @"C:\Docs\Result.html";

            // Load the template document
            Document doc = new Document(templatePath);

            // Create a data source object that matches the expressions used in the template.
            // Replace this with your actual data source.
            var dataSource = new SampleData
            {
                Title = "Report Title",
                Date = DateTime.Now,
                Value = 12345.67
            };

            // Populate the template with data using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the name used to reference the data source in the template.
            engine.BuildReport(doc, dataSource, "Data");

            // Configure HtmlFixedSaveOptions for fixed HTML output.
            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                // Ensure the format is set explicitly (optional, as Save will infer from options type)
                SaveFormat = SaveFormat.HtmlFixed,
                // Example: embed images as Base64 to keep a single HTML file.
                ExportEmbeddedImages = true,
                // Example: embed fonts if needed.
                ExportEmbeddedFonts = false,
                // Optimize the output by removing redundant canvases and merging glyphs.
                OptimizeOutput = true,
                // Show page borders (set false if not needed).
                ShowPageBorder = true
            };

            // Save the populated document as fixed HTML.
            doc.Save(outputPath, saveOptions);
        }

        // Sample data class – replace or extend with your actual data structure.
        public class SampleData
        {
            public string Title { get; set; }
            public DateTime Date { get; set; }
            public double Value { get; set; }
        }
    }
}
