using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportingExample
{
    // Simple data source class used in the template.
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
            // Path to the PDF template that contains LINQ Reporting tags, e.g. <<[ds.Title]>>.
            const string templatePath = @"C:\Templates\ReportTemplate.pdf";

            // Load the PDF template into an Aspose.Words Document.
            Document doc = new Document(templatePath); // load

            // Prepare the data source instance.
            var data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Amount = 123456.78m,
                Date = DateTime.Today
            };

            // Create the ReportingEngine and populate the template.
            ReportingEngine engine = new ReportingEngine(); // create
            // The data source name "ds" must match the name used in the template tags.
            engine.BuildReport(doc, data, "ds"); // build report

            // Configure PostScript save options.
            PsSaveOptions psOptions = new PsSaveOptions
            {
                // Ensure the format is set explicitly (optional, but clear).
                SaveFormat = SaveFormat.Ps
            };

            // Path for the resulting PostScript file.
            const string outputPath = @"C:\Output\ReportOutput.ps";

            // Save the populated document as a PostScript file.
            doc.Save(outputPath, psOptions); // save
        }
    }
}
