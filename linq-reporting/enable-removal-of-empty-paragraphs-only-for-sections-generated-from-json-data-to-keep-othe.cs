using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a static template (no data binding) with an empty paragraph.
        // -----------------------------------------------------------------
        string staticTemplatePath = Path.Combine(outputDir, "StaticTemplate.docx");
        Document staticTemplate = new Document();
        DocumentBuilder staticBuilder = new DocumentBuilder(staticTemplate);
        staticBuilder.Writeln("=== Static Section ===");
        staticBuilder.Writeln(""); // This empty paragraph should be preserved.
        staticBuilder.Writeln("End of static content.");
        staticTemplate.Save(staticTemplatePath);

        // -----------------------------------------------------------------
        // 2. Create a JSON-driven template with a foreach loop.
        // -----------------------------------------------------------------
        string jsonTemplatePath = Path.Combine(outputDir, "JsonTemplate.docx");
        Document jsonTemplate = new Document();
        DocumentBuilder jsonBuilder = new DocumentBuilder(jsonTemplate);
        jsonBuilder.Writeln("=== JSON Section ===");
        jsonBuilder.Writeln("<<foreach [item in Items]>>");
        jsonBuilder.Writeln("<<[item.Name]>>"); // May produce empty paragraph.
        jsonBuilder.Writeln("<</foreach>>");
        jsonBuilder.Writeln("End of JSON content.");
        jsonTemplate.Save(jsonTemplatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample JSON data.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alice" },
                new Item { Name = "" },      // This will generate an empty paragraph.
                new Item { Name = "Bob" },
                new Item { Name = null }    // Null also results in an empty paragraph.
            }
        };

        // -----------------------------------------------------------------
        // 4. Load the static template (no processing needed).
        // -----------------------------------------------------------------
        Document staticDoc = new Document(staticTemplatePath);

        // -----------------------------------------------------------------
        // 5. Load the JSON template and build the report with removal of empty paragraphs.
        // -----------------------------------------------------------------
        Document jsonDoc = new Document(jsonTemplatePath);
        ReportingEngine jsonEngine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        jsonEngine.BuildReport(jsonDoc, model, "model");

        // -----------------------------------------------------------------
        // 6. Append the processed JSON document to the static document.
        // -----------------------------------------------------------------
        staticDoc.AppendDocument(jsonDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 7. Save the final combined document.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "CombinedReport.docx");
        staticDoc.Save(resultPath, SaveFormat.Docx);
    }

    // Data model aligned with the JSON template.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        // Name may be null or empty; the engine will handle it.
        public string? Name { get; set; } = string.Empty;
    }
}
