using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a tag that calls the static TimeSpan.Parse method.
        // The engine will evaluate this expression and output the resulting TimeSpan.
        builder.Writeln("Parsed duration: <<[TimeSpan.Parse(\"01:30:00\")]>>");

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register System.TimeSpan so that static members (e.g., Parse) can be used in templates.
        engine.KnownTypes.Add(typeof(TimeSpan));

        // Build the report. No data source is required for this example, so we pass an empty object.
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
