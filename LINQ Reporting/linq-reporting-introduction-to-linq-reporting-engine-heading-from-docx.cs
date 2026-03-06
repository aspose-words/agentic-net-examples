using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDemo
{
    // Simple POCO class that will be used as the data source for the template.
    public class ReportData
    {
        public string Heading { get; set; }
        public string Description { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains LINQ Reporting Engine tags.
            // Example template content:
            //   <<[data.Heading]>>
            //   <<[data.Description]>>
            Document template = new Document("LinqReportingTemplate.docx");

            // Prepare the data source object.
            ReportData data = new ReportData
            {
                Heading = "LINQ Reporting Introduction",
                Description = "This document demonstrates how to use the Aspose.Words LINQ Reporting Engine."
            };

            // Create the reporting engine and populate the template with the data.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("data") matches the name used in the template tags.
            engine.BuildReport(template, data, "data");

            // Save the populated document.
            template.Save("LinqReportingResult.docx");
        }
    }
}
