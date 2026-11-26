using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Root data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Item model representing each row in the table.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public double Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data.
            string jsonPath = "data.json";
            string jsonContent = @"{
  ""Items"": [
    { ""Name"": ""Apple"",  ""Quantity"": 10, ""Price"": 0.5 },
    { ""Name"": ""Banana"", ""Quantity"": 5,  ""Price"": 0.3 },
    { ""Name"": ""Cherry"", ""Quantity"": 20, ""Price"": 0.2 }
  ]
}";
            File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

            // Deserialize JSON into the data model.
            ReportModel model = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonPath, Encoding.UTF8))!;

            // Create the template document programmatically.
            string templatePath = "Template.docx";
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Insert a title.
            builder.Writeln("Product Report");
            builder.Writeln();

            // Begin foreach loop over Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Start the table inside the loop.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Data row for each item.
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Quantity]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Price]>>");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            Document reportDoc = new(templatePath);

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
