using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- 1. Prepare sample JSON data ----------
        string jsonPath = "data.json";
        var jsonContent = @"{
  ""Categories"": [
    {
      ""Name"": ""Fruits"",
      ""Items"": [ ""Apple"", ""Banana"", ""Orange"" ]
    },
    {
      ""Name"": ""Vegetables"",
      ""Items"": [ ""Carrot"", ""Broccoli"" ]
    },
    {
      ""Name"": ""Beverages"",
      ""Items"": [ ""Coffee"", ""Tea"", ""Juice"" ]
    }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent);

        // ---------- 2. Create the data model that matches the JSON ----------
        var model = new ReportModel
        {
            Categories = new List<Category>
            {
                new Category { Name = "Fruits", Items = new List<string> { "Apple", "Banana", "Orange" } },
                new Category { Name = "Vegetables", Items = new List<string> { "Carrot", "Broccoli" } },
                new Category { Name = "Beverages", Items = new List<string> { "Coffee", "Tea", "Juice" } }
            }
        };

        // ---------- 3. Build the template document programmatically ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Create a bulleted list template and apply it to the builder
        List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // First level: categories
        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("<<[cat.Name]>>");

        // Second level: items inside each category
        builder.ListFormat.ListLevelNumber = 1; // indent for inner bullets
        builder.Writeln("<<foreach [item in cat.Items]>>");
        builder.Writeln("<<[item]>>");
        builder.Writeln("<</foreach>>");

        // Reset to outer level and close outer foreach
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before BuildReport)
        string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // ---------- 4. Load the template and generate the report ----------
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();

        // Use the wrapper name "model" in the template tags
        engine.BuildReport(loadedTemplate, model, "model");

        // ---------- 5. Save the final report ----------
        string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);

        // Inform that the process completed (no interactive input required)
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// ---------- Data model classes ----------
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
