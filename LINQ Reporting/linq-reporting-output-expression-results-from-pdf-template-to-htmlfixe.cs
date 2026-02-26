using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportingDemo
{
    // Simple data source class used in the template.
    public class InvoiceData
    {
        public string CustomerName { get; set; }
        public DateTime InvoiceDate { get; set; }
        public decimal TotalAmount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains LINQ Reporting tags, e.g. <<[data.CustomerName]>>.
            string templatePath = @"C:\Templates\InvoiceTemplate.pdf";

            // Load the PDF template into an Aspose.Words Document.
            Document doc = new Document(templatePath);

            // Prepare the data source instance.
            InvoiceData data = new InvoiceData
            {
                CustomerName = "Acme Corp.",
                InvoiceDate = DateTime.Today,
                TotalAmount = 1234.56m
            };

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the data source object.
            // The third parameter (optional) can be used to reference the object itself in the template.
            engine.BuildReport(doc, data, "data");

            // Configure HTML Fixed save options.
            HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
            {
                // Export images as external files (not embedded) for easier inspection.
                ExportEmbeddedImages = false,
                // Store resources (images, CSS, etc.) in a dedicated folder.
                ResourcesFolder = @"C:\Output\InvoiceResources",
                // Do not embed fonts to keep the output size small.
                ExportEmbeddedFonts = false,
                // Optional: pretty‑format the generated HTML for readability.
                PrettyFormat = true
            };

            // Ensure the resources folder exists.
            System.IO.Directory.CreateDirectory(htmlOptions.ResourcesFolder);

            // Save the populated document as HTML Fixed format.
            string outputPath = @"C:\Output\InvoiceReport.html";
            doc.Save(outputPath, htmlOptions);

            Console.WriteLine("Report generated successfully at: " + outputPath);
        }
    }
}
