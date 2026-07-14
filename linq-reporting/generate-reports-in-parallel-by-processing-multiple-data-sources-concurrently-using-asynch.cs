using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingParallel
{
    // Simple data item.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Root model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        // Entry point – async to allow awaiting parallel tasks.
        public static async Task Main(string[] args)
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Path for the shared template.
            string templatePath = Path.Combine(outputDir, "Template.docx");

            // 1. Create the LINQ Reporting template programmatically.
            CreateTemplate(templatePath);

            // 2. Prepare two distinct data sources.
            ReportModel model1 = new()
            {
                Title = "First Report",
                Items = new()
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" }
                }
            };

            ReportModel model2 = new()
            {
                Title = "Second Report",
                Items = new()
                {
                    new Item { Index = 1, Name = "Delta" },
                    new Item { Index = 2, Name = "Epsilon" },
                    new Item { Index = 3, Name = "Zeta" }
                }
            };

            // 3. Run report generation for each model in parallel.
            Task task1 = GenerateReportAsync(templatePath, model1, Path.Combine(outputDir, "Report1.docx"));
            Task task2 = GenerateReportAsync(templatePath, model2, Path.Combine(outputDir, "Report2.docx"));

            await Task.WhenAll(task1, task2);
        }

        // Creates a simple template containing a title and a foreach loop over Items.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title placeholder.
            builder.Writeln("Report: <<[model.Title]>>");
            builder.Writeln();

            // Begin foreach over Items.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item <<[item.Index]>> – <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(filePath);
        }

        // Loads the template, builds the report with the provided model, and saves the result.
        private static async Task GenerateReportAsync(string templatePath, ReportModel model, string outputPath)
        {
            await Task.Run(() =>
            {
                // Load the previously saved template.
                Document doc = new Document(templatePath);

                // Initialize the reporting engine.
                ReportingEngine engine = new ReportingEngine();

                // Build the report using the model as the root object named "model".
                engine.BuildReport(doc, model, "model");

                // Save the generated report.
                doc.Save(outputPath);
            });
        }
    }
}
