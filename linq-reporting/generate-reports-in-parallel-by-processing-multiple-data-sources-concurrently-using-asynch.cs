using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingParallel
{
    // Sample data model
    public class ReportData
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        private const string TemplateFileName = "Template.docx";
        private const string OutputFolder = "Reports";

        public static async Task Main()
        {
            // Register code page provider required by Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Ensure output folder exists
            Directory.CreateDirectory(OutputFolder);

            // Create the template document programmatically
            CreateTemplate();

            // Prepare multiple data sources
            var dataSources = new List<ReportData>
            {
                new()
                {
                    Persons = new()
                    {
                        new Person { Name = "Alice", Age = 30 },
                        new Person { Name = "Bob", Age = 25 },
                        new Person { Name = "Charlie", Age = 35 }
                    }
                },
                new()
                {
                    Persons = new()
                    {
                        new Person { Name = "David", Age = 28 },
                        new Person { Name = "Eve", Age = 22 },
                        new Person { Name = "Frank", Age = 40 }
                    }
                }
            };

            // Generate reports in parallel
            var tasks = new List<Task>();
            for (int i = 0; i < dataSources.Count; i++)
            {
                int index = i; // Capture loop variable
                tasks.Add(Task.Run(() => GenerateReportAsync(dataSources[index], index)));
            }

            await Task.WhenAll(tasks);
        }

        private static void CreateTemplate()
        {
            // Create a new blank document
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Write a simple heading
            builder.Writeln("Persons Report");
            builder.Writeln();

            // Insert LINQ Reporting tags
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            doc.Save(TemplateFileName);
        }

        private static async Task GenerateReportAsync(ReportData data, int reportIndex)
        {
            // Load the template
            var doc = new Document(TemplateFileName);

            // Build the report using Aspose.Words ReportingEngine
            var engine = new ReportingEngine();
            // BuildReport is synchronous; wrap it in Task.Run to avoid blocking
            await Task.Run(() => engine.BuildReport(doc, data, "model"));

            // Save the generated report
            string outputPath = Path.Combine(OutputFolder, $"Report_{reportIndex + 1}.docx");
            doc.Save(outputPath);
        }
    }
}
