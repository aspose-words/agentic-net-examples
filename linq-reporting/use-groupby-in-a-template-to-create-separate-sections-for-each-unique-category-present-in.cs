using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
    public int Value { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new List<Item>
        {
            new() { Category = "Fruits", Name = "Apple", Value = 10 },
            new() { Category = "Fruits", Name = "Banana", Value = 20 },
            new() { Category = "Vegetables", Name = "Carrot", Value = 15 },
            new() { Category = "Vegetables", Name = "Lettuce", Value = 5 },
            new() { Category = "Beverages", Name = "Coffee", Value = 30 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Create the LINQ Reporting template.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report grouped by Category");
        builder.Writeln();

        // Begin grouping loop.
        builder.Writeln("<<foreach [catGroup in Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[catGroup.Key]>>");
        builder.Writeln();

        // Items within each category.
        builder.Writeln("<<foreach [item in catGroup]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // Deserialize JSON into the model.
        var model = new ReportModel
        {
            Items = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new()
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        templateDoc.Save(outputPath);
    }
}
