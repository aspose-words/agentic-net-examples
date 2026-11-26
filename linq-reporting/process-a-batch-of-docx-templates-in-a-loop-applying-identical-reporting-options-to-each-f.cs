using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace BatchLinqReporting
{
    // Sample data model.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new()
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" }
                }
            };

            // Serialize sample data to JSON (demonstrates required package usage).
            string jsonSample = JsonConvert.SerializeObject(model, Formatting.Indented);
            File.WriteAllText("sampleData.json", jsonSample);

            // Create folders for templates and generated reports.
            string templatesDir = "Templates";
            string reportsDir = "Reports";
            Directory.CreateDirectory(templatesDir);
            Directory.CreateDirectory(reportsDir);

            // Create identical template files.
            const int templateCount = 3;
            for (int i = 1; i <= templateCount; i++)
            {
                string templatePath = Path.Combine(templatesDir, $"Template{i}.docx");
                CreateTemplate(templatePath);
            }

            // Process each template in the batch.
            foreach (string templatePath in Directory.GetFiles(templatesDir, "*.docx"))
            {
                // Load the template.
                Document doc = new(templatePath);

                // Configure the reporting engine.
                ReportingEngine engine = new();
                engine.Options = ReportBuildOptions.InlineErrorMessages;

                // Build the report.
                bool success = engine.BuildReport(doc, model, "model");

                // Save the generated report.
                string reportFileName = $"Report_{Path.GetFileName(templatePath)}";
                string reportPath = Path.Combine(reportsDir, reportFileName);
                doc.Save(reportPath);
            }
        }

        // Creates a simple LINQ Reporting template with a foreach loop.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new();
            DocumentBuilder builder = new(doc);

            builder.Writeln("Items Report");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }
    }
}
