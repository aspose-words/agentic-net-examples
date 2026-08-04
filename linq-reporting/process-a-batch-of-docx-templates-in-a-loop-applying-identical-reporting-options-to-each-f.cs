using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingBatch
{
    // Sample data model classes
    public class Order
    {
        public string CustomerName { get; set; } = "Acme Corp";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders
            string baseDir = Directory.GetCurrentDirectory();
            string templatesDir = Path.Combine(baseDir, "Templates");
            string outputDir = Path.Combine(baseDir, "Output");
            Directory.CreateDirectory(templatesDir);
            Directory.CreateDirectory(outputDir);

            // Create a few template files programmatically
            for (int i = 1; i <= 3; i++)
            {
                string templatePath = Path.Combine(templatesDir, $"Template{i}.docx");
                CreateTemplate(templatePath);
            }

            // Prepare a single data source that will be used for all reports
            Order sampleOrder = new Order
            {
                CustomerName = "Acme Corporation",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Widget A" },
                    new Item { Index = 2, Name = "Widget B" },
                    new Item { Index = 3, Name = "Widget C" }
                }
            };

            // Process each template in the folder
            foreach (string templateFile in Directory.GetFiles(templatesDir, "*.docx"))
            {
                // Load the template document
                Document doc = new Document(templateFile);

                // Configure the reporting engine
                ReportingEngine engine = new ReportingEngine();
                engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

                // Build the report using the same data source for every template
                engine.BuildReport(doc, sampleOrder, "order");

                // Save the generated report
                string outputFileName = Path.GetFileNameWithoutExtension(templateFile) + "_Report.docx";
                string outputPath = Path.Combine(outputDir, outputFileName);
                doc.Save(outputPath);
            }
        }

        // Creates a simple DOCX template with LINQ Reporting tags
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Header with customer name
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin foreach loop for items
            builder.Writeln("<<foreach [item in order.Items]>>");

            // Table with header and data rows
            Table table = builder.StartTable();

            // Header row
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Item Name");
            builder.EndRow();

            // Data row (repeated for each item)
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();

            // End of table and foreach block
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            doc.Save(filePath);
        }
    }
}
