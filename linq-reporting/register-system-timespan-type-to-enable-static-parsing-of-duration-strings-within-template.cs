using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that uses the static TimeSpan.Parse method.
        // Use double quotes for the string literal inside the tag.
        builder.Writeln("Duration: <<[TimeSpan.Parse(\"02:15:30\")]>>");

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register System.TimeSpan so that static members can be accessed from the template.
        engine.KnownTypes.Add(typeof(TimeSpan));

        // Build the report. No data source is required for this example,
        // so we pass an empty object as the root data source.
        engine.BuildReport(template, new object());

        // Save the generated document.
        template.Save("Report.docx");
    }
}
