// Load a PDF template, populate it with data using ReportingEngine, and save the result as PDF.
// The template can contain Aspose.Words reporting tags such as <<[Data.Name]>>.
// This example demonstrates the complete lifecycle: load → build report → save.

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsPdfTemplateExample
{
    // Sample data class that will be referenced from the template.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime GeneratedOn { get; set; }
        public decimal Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Paths to the input template PDF and the output PDF.
            string templatePath = @"C:\Templates\ReportTemplate.pdf";
            string outputPath   = @"C:\Output\ReportResult.pdf";

            // -----------------------------------------------------------------
            // 1. Load the PDF template into an Aspose.Words Document.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source that will be used by the ReportingEngine.
            // -----------------------------------------------------------------
            ReportData data = new ReportData
            {
                Title       = "Quarterly Sales Report",
                Author      = "John Doe",
                GeneratedOn = DateTime.Now,
                Amount      = 123456.78m
            };

            // -----------------------------------------------------------------
            // 3. Populate the template with the data.
            //    The data source name ("data") must match the name used in the template
            //    if the template references the source object itself (e.g., <<[data.Title]>>).
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "data");

            // -----------------------------------------------------------------
            // 4. Save the populated document as PDF.
            //    PdfSaveOptions can be customized if needed (e.g., preserve form fields,
            //    update fields, set page mode, etc.). Here we use the default options.
            // -----------------------------------------------------------------
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                // Example customizations (optional):
                // PreserveFormFields = true,
                // UpdateFields = true,
                // PageMode = PdfPageMode.UseOutlines
            };

            doc.Save(outputPath, saveOptions);
        }
    }
}
