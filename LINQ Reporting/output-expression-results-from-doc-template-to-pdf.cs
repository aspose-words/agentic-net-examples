using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Simple data class whose members are referenced in the Word template.
    public class ReportData
    {
        public string Title { get; set; }
        public DateTime ReportDate { get; set; }
        public double Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains expression tags like <<[ds.Title]>>.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path where the resulting PDF will be saved.
            string outputPdfPath = @"C:\Output\ReportResult.pdf";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source object.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                ReportDate = DateTime.Today,
                Amount = 123456.78
            };

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("ds") must match the name used in the template tags.
            engine.BuildReport(doc, data, "ds");

            // Configure PDF save options if needed (e.g., preserve document structure).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true   // optional: set to true to keep structure tags.
            };

            // Save the populated document as PDF.
            doc.Save(outputPdfPath, pdfOptions);
        }
    }
}
