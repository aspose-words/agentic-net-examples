using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        // Initialize non‑nullable reference types to avoid warnings.
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and add LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Template tags reference the root object named "model".
            builder.Writeln("Name: <<[model.Name]>>");
            builder.Writeln("Age: <<[model.Age]>>");

            // Prepare sample data.
            Person model = new Person
            {
                Name = "John Doe",
                Age = 30
            };

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report and capture the success flag.
            bool success = engine.BuildReport(doc, model, "model");

            // Output the result flag.
            Console.WriteLine($"Report build success: {success}");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }
    }
}
