using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that accesses DateTime.Now.
        builder.Writeln("Current date and time: <<[DateTime.Now]>>");

        // Set up the reporting engine and expose System.DateTime to the template.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report. No data source is required for this static call.
        engine.BuildReport(doc, new object());

        // Save the result.
        doc.Save("Report.docx");
    }
}
