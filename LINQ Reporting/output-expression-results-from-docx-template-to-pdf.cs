using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Simple data source class whose members are referenced in the DOCX template.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime Created { get; set; }
        public decimal Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains expression tags like <<[data.Title]>>.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path where the resulting PDF will be saved.
            string pdfOutputPath = @"C:\Output\ReportResult.pdf";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source.
            var data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Author = "John Doe",
                Created = DateTime.Now,
                Amount = 123456.78m
            };

            // Populate the template with data using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("data") must match the name used in the template tags.
            engine.BuildReport(doc, data, "data");

            // Configure PDF save options (e.g., normal color rendering).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ColorMode = ColorMode.Normal
            };

            // Save the populated document as PDF.
            doc.Save(pdfOutputPath, pdfOptions);
        }
    }
}
