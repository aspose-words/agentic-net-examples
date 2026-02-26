using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingExample
{
    // Simple data source class whose members will be referenced from the template.
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
            // Load the PDF template that contains LINQ Reporting tags (e.g. <<[data.Title]>>).
            Document template = new Document("Template.pdf");

            // Prepare the data source.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Amount = 123456.78m,
                Date = DateTime.Today
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members to be ignored (optional).
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "N/A"
            };

            // Populate the template with the data source.
            // The third argument is the name used to reference the data source inside the template.
            engine.BuildReport(template, data, "data");

            // Configure PDF save options if needed.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: preserve form fields as form fields in the PDF.
                PreserveFormFields = true
            };

            // Save the populated document as a PDF file.
            template.Save("Result.pdf", pdfOptions);
        }
    }
}
