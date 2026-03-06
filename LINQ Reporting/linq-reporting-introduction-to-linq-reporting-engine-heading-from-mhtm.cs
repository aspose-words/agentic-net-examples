using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    class Program
    {
        static void Main()
        {
            // Load the MHTML template that contains the heading "LINQ Reporting Introduction to LINQ Reporting Engine".
            Document template = new Document("Template.mhtml");

            // Create an instance of the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. No data source is required for a static heading,
            // so we pass an empty object and a null data source name.
            engine.BuildReport(template, new object(), null);

            // Save the populated document to DOCX format.
            template.Save("LinqReportingOutput.docx");
        }
    }
}
