using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ParallelLinqReporting
{
    // Sample data model
    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        private const string TemplateFileName = "template.docx";
        private const string OutputFolder = "output";

        public static async Task Main()
        {
            // Ensure output directory exists
            Directory.CreateDirectory(OutputFolder);

            // Create and save the template document
            CreateTemplate(TemplateFileName);

            // Prepare multiple data sources
            var models = new List<ReportModel>
            {
                new()
                {
                    Title = "Team Alpha Report",
                    Persons = new()
                    {
                        new() { Name = "Alice", Age = 30 },
                        new() { Name = "Bob", Age = 25 }
                    }
                },
                new()
                {
                    Title = "Team Beta Report",
                    Persons = new()
                    {
                        new() { Name = "Charlie", Age = 28 },
                        new() { Name = "Diana", Age = 32 },
                        new() { Name = "Eve", Age = 27 }
                    }
                }
            };

            // Generate reports in parallel
            var tasks = new List<Task>();
            for (int i = 0; i < models.Count; i++)
            {
                int index = i; // Capture loop variable
                string outputPath = Path.Combine(OutputFolder, $"Report_{index + 1}.docx");
                tasks.Add(GenerateReportAsync(TemplateFileName, models[index], outputPath));
            }

            await Task.WhenAll(tasks);
        }

        // Creates the LINQ Reporting template programmatically
        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<<foreach [p in model.Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }

        // Generates a single report based on the provided model
        private static Task GenerateReportAsync(string templatePath, ReportModel model, string outputPath)
        {
            return Task.Run(() =>
            {
                // Load the template
                var doc = new Document(templatePath);

                // Build the report
                var engine = new ReportingEngine();
                bool success = engine.BuildReport(doc, model, "model");

                // Save the generated document
                if (success)
                {
                    doc.Save(outputPath);
                }
            });
        }
    }
}
