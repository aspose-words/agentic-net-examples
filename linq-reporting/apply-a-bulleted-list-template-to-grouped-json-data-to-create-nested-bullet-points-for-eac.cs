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
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Create sample JSON data ----------
        string json = @"
[
  {
    ""Name"": ""Fruits"",
    ""Items"": [
      { ""Name"": ""Apple"" },
      { ""Name"": ""Banana"" },
      { ""Name"": ""Orange"" }
    ]
  },
  {
    ""Name"": ""Vegetables"",
    ""Items"": [
      { ""Name"": ""Carrot"" },
      { ""Name"": ""Broccoli"" },
      { ""Name"": ""Spinach"" }
    ]
  }
]";
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, json);

        // Deserialize JSON into model objects.
        List<Category> categories = JsonConvert.DeserializeObject<List<Category>>(File.ReadAllText(jsonPath)) ?? new();
        var reportData = new ReportData { Categories = categories };

        // ---------- Create the LINQ Reporting template ----------
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Outer foreach: categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        // Category name as a heading (bold).
        builder.Font.Bold = true;
        builder.Writeln("<<[cat.Name]>>");
        builder.Font.Bold = false;

        // Inner foreach: items within a category.
        builder.Writeln("<<foreach [item in cat.Items]>>");
        // Apply bullet list formatting for each item.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("<<[item.Name]>>");
        // Reset list formatting after the line.
        builder.ListFormat.RemoveNumbers();
        builder.Writeln("<</foreach>>"); // End inner foreach.

        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save the template.
        templateDoc.Save(templatePath);

        // ---------- Load the template and build the report ----------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, reportData, "data");

        // Save the generated report.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// ---------- Data model ----------
public class ReportData
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
}
