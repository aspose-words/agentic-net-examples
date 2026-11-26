using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace LinqReportingJsonArray
{
    // Root model that matches the JSON structure.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register the code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create sample JSON data and write it to a local file.
            // -----------------------------------------------------------------
            const string jsonPath = "data.json";
            var sampleData = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Title = "Item A", Description = "First item description.", Amount = 123.45m },
                    new Item { Title = "Item B", Description = "Second item description.", Amount = 678.90m },
                    new Item { Title = "Item C", Description = "Third item description.", Amount = 42.00m }
                }
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

            // -----------------------------------------------------------------
            // 2. Build the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Report title.
            builder.Writeln("Report generated from JSON array");
            builder.Writeln();

            // Begin the foreach block that will repeat the data rows.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table for each iteration (header row + data row).
            Table table = builder.StartTable();

            // Header row (appears for every item – acceptable for a simple example).
            builder.InsertCell();
            builder.Writeln("Title");
            builder.InsertCell();
            builder.Writeln("Description");
            builder.InsertCell();
            builder.Writeln("Amount");
            builder.EndRow();

            // Data row – values are taken from the current item.
            builder.InsertCell();
            builder.Write("<<[item.Title]>>");
            builder.InsertCell();
            builder.Write("<<[item.Description]>>");
            builder.InsertCell();
            builder.Write("<<[item.Amount]>>");
            builder.EndRow();

            // End of the table for this iteration.
            builder.EndTable();

            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and generate the final report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // Load JSON data into the model.
            string jsonContent = File.ReadAllText(jsonPath);
            var model = JsonConvert.DeserializeObject<ReportModel>(jsonContent) ?? new ReportModel();

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
