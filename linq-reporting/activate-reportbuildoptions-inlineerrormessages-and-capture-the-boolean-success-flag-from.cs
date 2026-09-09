using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample data object.
            var person = new Person
            {
                Name = "John Doe",
                Age = 30
            };

            // Create a new blank document and insert a LINQ Reporting tag.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            // The tag references the root object name "person".
            builder.Writeln("<<[person.Name]>> is <<[person.Age]>> years old.");

            // Configure the reporting engine to inline error messages.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report and capture the success flag.
            bool success = engine.BuildReport(doc, person, "person");

            // Save the resulting document.
            doc.Save("ReportOutput.docx");

            // Output the success flag (no interactive input required).
            Console.WriteLine($"Report build success: {success}");
        }
    }
}
