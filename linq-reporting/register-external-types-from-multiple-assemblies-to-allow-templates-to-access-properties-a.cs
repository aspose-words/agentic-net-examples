using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simulate a type from an external assembly.
    public static class Utils
    {
        public static string GetGreeting(string name) => $"Hello, {name}!";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank template document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags that use static members from external types.
            // Correct syntax uses a dot (.) to separate type and member.
            builder.Writeln("Value of PI: <<[Math.PI]>>");
            builder.Writeln("Custom greeting: <<[Utils.GetGreeting(\"World\")]>>");

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register external types so the template can reference them without full qualification.
            engine.KnownTypes.Add(typeof(System.Math));
            engine.KnownTypes.Add(typeof(Utils));

            // Build the report. No data source is required for this example.
            engine.BuildReport(doc, new object(), "");

            // Save the generated document.
            doc.Save("ReportOutput.docx");
        }
    }
}
