using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string json = @"{
  ""CustomerName"": ""Acme Corp"",
  ""Items"": [
    { ""Index"": 1, ""Name"": ""Widget"", ""Quantity"": 10 },
    { ""Index"": 2, ""Name"": ""Gadget"", ""Quantity"": 5 },
    { ""Index"": 3, ""Name"": ""Doohickey"", ""Quantity"": 7 }
  ]
}";
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, json, Encoding.UTF8);

        // Deserialize JSON into model.
        Order? order = JsonConvert.DeserializeObject<Order>(File.ReadAllText(jsonPath, Encoding.UTF8));
        if (order == null)
        {
            Console.WriteLine("Failed to deserialize JSON.");
            return;
        }

        // Create Word template with LINQ Reporting tags.
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Customer line.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Header table (static, not repeated).
        Table headerTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Index");
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.EndRow();
        builder.EndTable();

        // Begin foreach loop for items.
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Data rows table (repeated for each item).
        Table dataTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("<<[item.Index]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}
