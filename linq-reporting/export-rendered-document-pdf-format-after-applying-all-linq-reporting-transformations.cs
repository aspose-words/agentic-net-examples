using System;
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
            // Path to the template document that contains LINQ Reporting tags.
            const string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source that will be used by the ReportingEngine.
            var data = new ReportData
            {
                Title = "Annual Report",
                Author = "John Doe",
                Year = DateTime.Now.Year
            };

            // Create the reporting engine and populate the template.
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the name used to reference the data source inside the template.
            engine.BuildReport(doc, data, "data");

            // Export the fully populated document to PDF.
            // The file extension determines the save format (PDF) automatically.
            const string outputPdfPath = @"C:\Docs\Report.pdf";
            doc.Save(outputPdfPath);
        }
    }
}
