using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that accesses DateTime.Now.
        builder.Writeln("Current date and time: <<[DateTime.Now]>>");

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Add System.DateTime to the known types collection.
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report. No data source is needed for this static call.
        engine.BuildReport(doc, new object(), "");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
