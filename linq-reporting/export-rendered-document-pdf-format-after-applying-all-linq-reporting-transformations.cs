using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsPdfExport
{
    // Sample data source class – replace with your actual data model.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public int Year { get; set; }
        // Add other properties that are referenced in the template.
    }

    class Program
    {
        static void Main()
        {
            // Create a simple in‑memory template document with LINQ Reporting tags.
            Document doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Report Title: <<[data.Title]>>");
            builder.Writeln("Author: <<[data.Author]>>");
            builder.Writeln("Year: <<[data.Year]>>");

            // Prepare the data source that will be used by the ReportingEngine.
            var data = new ReportData
            {
                Title = "Annual Report",
                Author = "John Doe",
                Year = DateTime.Now.Year
            };

            // Build the report using the data source.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, data, "data");

            // Export the fully populated document to PDF in the current directory.
            string outputPdfPath = Path.Combine(Environment.CurrentDirectory, "Report.pdf");
            doc.Save(outputPdfPath);
            Console.WriteLine($"Report saved to: {outputPdfPath}");
        }
    }
}
