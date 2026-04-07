using System;
using System.Collections.Generic;
using System.Globalization;
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
        string jsonContent = @"{
  ""Items"": [
    { ""Name"": ""Widget"", ""Quantity"": 3, ""UnitPrice"": 19.99 },
    { ""Name"": ""Gadget"", ""Quantity"": 5, ""UnitPrice"": 9.50 },
    { ""Name"": ""Doohickey"", ""Quantity"": 2, ""UnitPrice"": 24.75 }
  ]
}";
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, jsonContent);

        // Deserialize JSON into model.
        RootModel root = JsonConvert.DeserializeObject<RootModel>(File.ReadAllText(jsonPath))!;

        // Create template document programmatically.
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Invoice Report");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create table header.
        Table table = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Item");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.InsertCell(); builder.Writeln("Unit Price");
        builder.InsertCell(); builder.Writeln("Total");
        builder.EndRow();

        // Row for each item.
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.UnitPrice.ToString(\"C\")]>>");
        builder.InsertCell(); builder.Writeln("<<[(item.Quantity * item.UnitPrice).ToString(\"C\")]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, root, "root");

        // Save the generated report.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);
        string reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);

        // Inform the user (no waiting for input).
        Console.WriteLine($"Report generated at: {Path.GetFullPath(reportPath)}");
    }
}

// Root wrapper model.
public class RootModel
{
    public List<ItemModel> Items { get; set; } = new();
}

// Item model.
public class ItemModel
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
