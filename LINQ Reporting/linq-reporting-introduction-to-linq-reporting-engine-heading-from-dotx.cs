using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTX template that contains the heading <<[report.Title]>>.
        Document template = new Document("Template.dotx");

        // Prepare a simple data source with the values to fill the template.
        var data = new ReportData
        {
            Title = "LINQ Reporting Introduction",
            Description = "This document demonstrates how to use the LINQ Reporting Engine with a DOTX template."
        };

        // Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data, "report");

        // Save the generated report.
        template.Save("ReportOutput.docx");
    }

    // POCO class used as the data source for the report.
    public class ReportData
    {
        public string Title { get; set; }
        public string Description { get; set; }
    }
}
