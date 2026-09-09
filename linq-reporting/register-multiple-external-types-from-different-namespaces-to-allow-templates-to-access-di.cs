using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Custom helper class placed in a separate namespace.
    // The static method will be accessed from the template via ReportingEngine.KnownTypes.
    namespace MyNamespace
    {
        public static class CustomHelper
        {
            public static string GetMessage()
            {
                return "Hello from CustomHelper!";
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags that reference static members from different namespaces.
            builder.Writeln("Value of Math.PI: <<[Math.PI]>>");
            builder.Writeln("Custom message: <<[MyNamespace.CustomHelper.GetMessage()]>>");

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register external types so that the template can access their static members.
            engine.KnownTypes.Add(typeof(System.Math));
            engine.KnownTypes.Add(typeof(MyNamespace.CustomHelper));

            // Build the report. No root data object is required for this example,
            // so we pass an empty object and an empty data source name.
            engine.BuildReport(doc, new object(), string.Empty);

            // Save the generated document.
            doc.Save("Report.docx");
        }
    }
}
