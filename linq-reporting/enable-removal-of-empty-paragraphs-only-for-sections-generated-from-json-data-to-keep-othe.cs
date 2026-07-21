using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Reporting;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // 2. Create sample JSON data file.
        string jsonPath = Path.Combine(outputDir, "Data.json");
        CreateJsonData(jsonPath);

        // 3. Load the template document.
        Document doc = new Document(templatePath);

        // 4. Load JSON data source.
        using (FileStream jsonStream = File.OpenRead(jsonPath))
        {
            JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

            // 5. Configure the reporting engine to remove empty paragraphs.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // 6. Build the report. The root object name is "items" to match the template tags.
            engine.BuildReport(doc, jsonDataSource, "items");
        }

        // 7. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }

    // Creates a template with a static section and a JSON‑driven section.
    private static void CreateTemplate(string filePath)
    {
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // ----- Static section (should remain untouched) -----
        builder.Writeln("=== Static Section ===");
        builder.Writeln("This paragraph is always kept, even if empty.");
        builder.Writeln(""); // an empty paragraph that we do NOT want removed

        // Start a new section for JSON‑generated content.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("=== JSON Section ===");

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        // The exclamation mark after the tag marks this paragraph for selective removal.
        builder.Writeln("Description: <<[item.Description]>>!");
        builder.Writeln("<</foreach>>");

        template.Save(filePath);
    }

    // Generates a JSON file with some empty / missing values.
    private static void CreateJsonData(string filePath)
    {
        List<Item> items = new()
        {
            new Item { Name = "Apple",  Description = "Fresh red apple" },
            new Item { Name = "Banana", Description = "" },               // empty description
            new Item { Name = "Cherry" }                                 // missing description
        };

        string json = JsonSerializer.Serialize(items);
        File.WriteAllText(filePath, json);
    }

    // Simple data model matching the JSON structure.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
    }
}
