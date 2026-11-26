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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // 2. Create sample JSON data.
        string jsonPath = Path.Combine(outputDir, "Data.json");
        CreateSampleJson(jsonPath);

        // 3. Load JSON into the model.
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonPath))!;

        // 4. Load the template.
        Document doc = new(templatePath);

        // 5. Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // 6. Build the report.
        engine.BuildReport(doc, model, "model");

        // 7. Save the generated report.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Static content – will remain untouched.
        builder.Writeln("=== Report Title ===");
        builder.Writeln("This paragraph is static and will stay even if empty.");

        // Section generated from JSON data.
        builder.Writeln("=== Items Section ===");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Description]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    private static void CreateSampleJson(string path)
    {
        var sample = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Description = "First item description." },
                new Item { Description = "" }, // Empty description – should be removed.
                new Item { Description = "Third item description." },
                new Item { Description = null } // Null description – also results in empty paragraph.
            }
        };

        string json = JsonConvert.SerializeObject(sample, Formatting.Indented);
        File.WriteAllText(path, json);
    }
}

// Root model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the JSON data.
public class Item
{
    public string? Description { get; set; }
}
