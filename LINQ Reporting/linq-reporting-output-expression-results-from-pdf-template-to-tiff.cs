using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingToTiff
{
    // Simple data source class used in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public DateTime ReportDate { get; set; }
        public decimal Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains LINQ Reporting tags, e.g. <<[ds.Title]>>.
            string templatePath = @"C:\Templates\ReportTemplate.pdf";

            // Load the PDF template into a Document object.
            Document doc = new Document(templatePath);

            // Prepare the data source instance.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                ReportDate = DateTime.Today,
                Amount = 123456.78m
            };

            // Create the ReportingEngine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "ds" must match the name used in the template tags.
            engine.BuildReport(doc, data, "ds");

            // Save the populated document as a TIFF image.
            // Each page of the document will be saved as a separate TIFF frame.
            string outputPath = @"C:\Output\Report.tiff";
            doc.Save(outputPath, SaveFormat.Tiff);
        }
    }
}
