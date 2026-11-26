using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new Root
        {
            Categories = new List<Category>
            {
                new Category
                {
                    Name = "Fruits",
                    Items = new List<string> { "Apple", "Banana", "Orange" }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new List<string> { "Carrot", "Lettuce", "Tomato" }
                }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Create the LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Category Report");
        builder.Writeln();

        // Outer foreach for categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        // Category name as a heading.
        builder.Writeln("<<[cat.Name]>>");
        // Inner foreach for items.
        builder.Writeln("<<foreach [item in cat.Items]>>");
        // Bullet point for each item.
        builder.Writeln("• <<[item]>>");
        builder.Writeln("<</foreach>>");
        // End of outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Load JSON data into the model.
        var jsonContent = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<Root>(jsonContent)!;

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "root");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model.
public class Root
{
    public List<Category> Categories { get; set; } = new();
}

// Category model.
public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
