using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOTM template that contains the heading placeholder.
            // The template file should exist at the specified path.
            Document template = new Document("Template.dotm");

            // Create an instance of the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. In this simple example we do not need any data source,
            // so we pass an empty anonymous object.
            engine.BuildReport(template, new { });

            // Save the generated document as DOCX.
            template.Save("LinqReportingResult.docx");
        }
    }
}
