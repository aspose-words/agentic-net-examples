using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains LINQ Reporting tags, e.g. <<[report.Title]>> and <<[report.Total]>>.
        Document doc = new Document("Template.docx");

        // Create a simple data source object that will be referenced in the template.
        var data = new ReportData
        {
            Title = "Quarterly Sales Report",
            Total = 12345.67M
        };

        // Populate the template with the data source using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "report"); // "report" is the name used in the template.

        // Save the populated document directly to TIFF format.
        doc.Save("Report.tiff"); // Extension determines the SaveFormat (TIFF).
    }

    // Simple POCO class used as the data source for the report.
    public class ReportData
    {
        public string Title { get; set; }
        public decimal Total { get; set; }
    }
}
