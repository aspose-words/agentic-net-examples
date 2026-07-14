using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with only a Name property.
    public class Person
    {
        public Person(string name) => Name = name;
        public string Name { get; set; } = string.Empty; // Initialized to avoid nullable warnings.
        // Note: Age property is intentionally missing to demonstrate missing member handling.
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags. The Age tag does not exist in the Person class.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>"); // Missing member.

            // Prepare the data source.
            Person person = new Person("John Doe");

            // Configure the reporting engine to treat missing members as null.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.MissingMemberMessage = "N/A"; // Optional custom message for missing members.

            // Build the report. The root object name must match the name used in the template tags.
            engine.BuildReport(doc, person, "person");

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the generated report.
            string outputPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(outputPath);

            // Indicate completion (no interactive input).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
