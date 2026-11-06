using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSingleRowList
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                Persons = new()
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 28 }
                }
            };

            // Create template document programmatically
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a single‑row list template using LINQ Reporting tags
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before building the report)
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Load the template for report generation
            var doc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Optionally indicate success (no interactive input)
            Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
        }
    }
}
