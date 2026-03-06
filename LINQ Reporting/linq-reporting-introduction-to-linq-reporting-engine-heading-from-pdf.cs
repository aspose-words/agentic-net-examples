using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the PDF template that contains the placeholder for the heading.
        Document doc = new Document("Template.pdf");

        // Simple data source with a property that matches the placeholder in the template.
        var data = new ReportData
        {
            Title = "LINQ Reporting Introduction to LINQ Reporting Engine"
        };

        // Populate the template using the LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "report");

        // Save the generated report as a DOCX file.
        doc.Save("ReportOutput.docx");
    }

    // Data source class used by the reporting engine.
    public class ReportData
    {
        public string Title { get; set; }
    }
}
