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

        // Sample JSON data.
        string json = @"{
  ""Items"": [
    { ""Name"": ""Apple"",  ""Price"": 0.5, ""Quantity"": 10 },
    { ""Name"": ""Banana"", ""Price"": 0.3, ""Quantity"": 5 },
    { ""Name"": ""Orange"", ""Price"": 0.8, ""Quantity"": 7 }
  ]
}";
        // Deserialize JSON into model.
        Order order = JsonConvert.DeserializeObject<Order>(json)!;

        // Create a Word template programmatically.
        const string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Items Report");
        builder.Writeln();

        // Header table (static, not repeated).
        Table headerTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Price");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.InsertCell(); builder.Writeln("Total (Price * Quantity)");
        builder.EndRow();
        builder.EndTable();

        // Data rows – wrapped in a foreach block.
        builder.Writeln("<<foreach [item in Items]>>");
        Table dataTable = builder.StartTable();
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price * item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the root object name "order".
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes.
public class Order
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}
