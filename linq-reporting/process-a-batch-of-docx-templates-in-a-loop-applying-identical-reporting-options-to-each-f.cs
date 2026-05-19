using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBatch
{
    public class Program
    {
        public static void Main()
        {
            // Prepare folders.
            string baseDir = Directory.GetCurrentDirectory();
            string templatesDir = Path.Combine(baseDir, "Templates");
            string reportsDir = Path.Combine(baseDir, "Reports");
            Directory.CreateDirectory(templatesDir);
            Directory.CreateDirectory(reportsDir);

            // Create a few identical template files.
            for (int i = 1; i <= 3; i++)
            {
                string templatePath = Path.Combine(templatesDir, $"Template{i}.docx");
                CreateTemplate(templatePath);
            }

            // Prepare a sample data model.
            ReportModel model = new ReportModel
            {
                CustomerName = "Acme Corp",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Widget" },
                    new Item { Index = 2, Name = "Gadget" },
                    new Item { Index = 3, Name = "Doohickey" }
                }
            };

            // Configure the reporting engine once.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Process each template in the batch.
            foreach (string templateFile in Directory.GetFiles(templatesDir, "*.docx"))
            {
                // Load the template.
                Document doc = new Document(templateFile);

                // Build the report using the same model and root name "model".
                engine.BuildReport(doc, model, "model");

                // Save the generated report.
                string reportFileName = "Report_" + Path.GetFileName(templateFile);
                string reportPath = Path.Combine(reportsDir, reportFileName);
                doc.Save(reportPath);
            }
        }

        // Creates a simple DOCX template with LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Customer: <<[model.CustomerName]>>");
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }
    }

    // Root data model for the report.
    public class ReportModel
    {
        public string CustomerName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    // Item used in the collection.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }
}
