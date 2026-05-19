using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create sample JSON data.
        string jsonPath = Path.Combine(workDir, "data.json");
        var sampleData = new
        {
            Items = new List<Item>
            {
                new Item { Category = "Fruits", Name = "Apple", Value = 10 },
                new Item { Category = "Fruits", Name = "Banana", Value = 20 },
                new Item { Category = "Vegetables", Name = "Carrot", Value = 15 },
                new Item { Category = "Fruits", Name = "Orange", Value = 12 },
                new Item { Category = "Vegetables", Name = "Lettuce", Value = 8 }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Load JSON into the model.
        string jsonContent = File.ReadAllText(jsonPath);
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(jsonContent) ?? new ReportModel();

        // 3. Create the LINQ Reporting template.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin grouping by Category using LINQ GroupBy inside the foreach tag.
        builder.Writeln("<<foreach [g in Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[g.Key]>>");
        builder.Writeln("<<foreach [i in g]>>");
        builder.Writeln("- <<[i.Name]>>: <<[i.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 4. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 5. Build the report using the model as the root data source named "model".
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the generated report.
        string outputPath = Path.Combine(workDir, "report.docx");
        reportDoc.Save(outputPath);
    }
}

// Root model class.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Data item class.
public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
