using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new empty document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a placeholder that uses the static TimeSpan.Parse method.
        // The ReportingEngine will evaluate this expression when building the report.
        builder.Writeln("Duration: {{TimeSpan.Parse(\"01:30:00\")}}");

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register System.TimeSpan so that its static members can be used in templates.
        engine.KnownTypes.Add(typeof(TimeSpan));

        // Build the report. No external data source is required for this example,
        // so we pass an empty anonymous object.
        engine.BuildReport(doc, new { });

        // Save the generated document.
        doc.Save("Output.docx");
    }
}
