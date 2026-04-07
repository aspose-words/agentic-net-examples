using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingInlineErrorDemo
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Person> Persons { get; set; } = new();

        // Additional property to demonstrate direct root access (optional).
        public string Title { get; set; } = "Person List";
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Write a title placeholder.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln(); // Empty line for readability.

            // Write a foreach block that iterates over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template document (could also reuse the same instance).
            Document doc = new Document(templatePath);

            // 3. Prepare sample data.
            ReportModel model = new()
            {
                Title = "Sample Persons",
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 45 },
                    new Person { Name = "Charlie", Age = 28 }
                }
            };

            // 4. Configure the ReportingEngine with InlineErrorMessages option.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // 5. Build the report and capture the success flag.
            bool success = engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            // 7. Output the success flag (no interactive input required).
            Console.WriteLine($"Report build success: {success}");
        }
    }
}
