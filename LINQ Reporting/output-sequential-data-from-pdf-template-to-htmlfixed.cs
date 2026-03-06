using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace PdfToHtmlFixedExample
{
    // Simple data class that will be used as a data source for the reporting engine.
    public class ReportData
    {
        public string Title { get; set; }
        public List<Item> Items { get; set; }

        public class Item
        {
            public int Index { get; set; }
            public string Description { get; set; }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains reporting tags (e.g. <<[ds.Title]>>, <<foreach [ds.Items]>><<[Index]>> - <<[Description]>><</foreach>>)
            string pdfTemplatePath = @"C:\Templates\ReportTemplate.pdf";

            // Load the PDF template into an Aspose.Words Document.
            Document doc = new Document(pdfTemplatePath);

            // Prepare the data that will be merged into the template.
            var data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Items = new List<ReportData.Item>
                {
                    new ReportData.Item { Index = 1, Description = "North America" },
                    new ReportData.Item { Index = 2, Description = "Europe" },
                    new ReportData.Item { Index = 3, Description = "Asia" }
                }
            };

            // Use the ReportingEngine to populate the template with the data.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "ds" must match the name used in the template tags.
            engine.BuildReport(doc, data, "ds");

            // Configure HtmlFixedSaveOptions.
            HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
            {
                // Export form fields as interactive HTML elements (if any exist in the template).
                ExportFormFields = true,

                // Optimize output to remove redundant canvases and improve file size.
                OptimizeOutput = true,

                // Do not embed images as Base64; keep them as separate files for easier inspection.
                ExportEmbeddedImages = false,

                // Specify a folder where external resources (images, CSS, fonts) will be saved.
                ResourcesFolder = Path.Combine(Environment.CurrentDirectory, "HtmlResources")
            };

            // Ensure the resources folder exists.
            Directory.CreateDirectory(htmlOptions.ResourcesFolder);

            // Save the populated document as fixed HTML.
            string outputHtmlPath = Path.Combine(Environment.CurrentDirectory, "ReportOutput.html");
            doc.Save(outputHtmlPath, htmlOptions);

            Console.WriteLine("Document has been converted to fixed HTML successfully.");
        }
    }
}
