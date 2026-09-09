using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Custom external type whose static members can be accessed from the template.
    public static class CustomHelper
    {
        // Static property that will be read by the LINQ Reporting engine.
        public static string Greeting => "Hello from the registered external type!";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting tag that references the static property of CustomHelper.
            // The type name is resolved because we will register it in KnownTypes.
            builder.Writeln("<<[CustomHelper.Greeting]>>");

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register the custom external type so that the template can use it without reflection.
            engine.KnownTypes.Add(typeof(CustomHelper));

            // Build the report. No data source is required for this example.
            engine.BuildReport(doc, new object());

            // Save the generated document.
            doc.Save("Report.docx");
        }
    }
}
