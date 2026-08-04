using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingAsync
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        // Async entry point.
        public static async Task Main(string[] args)
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a template document with LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Add a heading.
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin a foreach loop over the Persons collection.
            builder.Writeln("<<foreach [person in Persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before loading for the engine).
            string templatePath = Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // 2. Load the template back (simulating a real-world scenario where the template is a file).
            Document loadedTemplate = new Document(templatePath);

            // 3. Prepare sample data.
            ReportModel model = new()
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 45 },
                    new Person { Name = "Charlie", Age = 28 }
                }
            };

            // 4. Build the report asynchronously to avoid blocking the calling thread.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            bool success = await Task.Run(() => engine.BuildReport(loadedTemplate, model, "model"));

            // 5. Save the generated report.
            string reportPath = Path.Combine(outputDir, "ReportOutput.docx");
            loadedTemplate.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to: {reportPath}");
        }
    }
}
