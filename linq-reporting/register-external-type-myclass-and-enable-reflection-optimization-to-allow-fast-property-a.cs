using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample external type to be used in the template.
    public class MyClass
    {
        // Static property accessed from the template.
        public static string Value { get; } = "42";

        // Static method accessed from the template.
        public static string GetMessage()
        {
            return "Hello from MyClass!";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Enable reflection optimization for faster property access.
            ReportingEngine.UseReflectionOptimization = true;

            // Create a simple template document with LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Value: <<[MyClass.Value]>>");
            builder.Writeln("Message: <<[MyClass.GetMessage()]>>");

            // Set up the reporting engine and register the external type.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(MyClass));

            // Build the report. No data source is required because the template uses only static members.
            engine.BuildReport(doc, new object());

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
