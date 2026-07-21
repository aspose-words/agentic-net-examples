using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Existing property that will be displayed correctly.
        public string ExistingName { get; set; } = "John Doe";

        // Collection property that will be iterated in the template.
        public Person[] People { get; set; } = new[]
        {
            new Person { Id = 1, Name = "Alice" },
            new Person { Id = 2, Name = "Bob" }
        };
    }

    public class Person
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare the output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a new blank document and a builder to insert template tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Normal tag – will be replaced with the actual value.
            builder.Writeln("Existing name: <<[model.ExistingName]>>");

            // Tag referencing a missing object – will be treated as null.
            builder.Writeln("Missing object name: <<[model.MissingObject.Name]>>");

            // Foreach loop over a missing collection – will be treated as empty.
            builder.Writeln("Missing collection loop:");
            builder.Writeln("<<foreach [item in model.MissingCollection]>>");
            builder.Writeln("Item Id: <<[item.Id]>>");
            builder.Writeln("<</foreach>>");

            // Loop over an existing collection to show normal behavior.
            builder.Writeln("People:");
            builder.Writeln("<<foreach [person in model.People]>>");
            builder.Writeln("Id: <<[person.Id]>>, Name: <<[person.Name]>>");
            builder.Writeln("<</foreach>>");

            // Initialize the reporting engine with the required options.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages,
                MissingMemberMessage = "N/A"
            };

            // Build the report using the document, the data model, and the root name "model".
            bool success = engine.BuildReport(doc, new ReportModel(), "model");

            // Save the resulting document.
            string outputPath = Path.Combine(outputDir, "ReportWithMissingMembers.docx");
            doc.Save(outputPath);

            // Output simple status information (no interactive prompts).
            Console.WriteLine($"Report generation success flag: {success}");
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
