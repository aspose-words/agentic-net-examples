using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create the template document.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Alpha" },
                new Item { Name = "Beta" },
                new Item { Name = "Gamma" },
                new Item { Name = "Delta" }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Custom Font Styling Example");
        builder.Writeln();

        // Begin foreach over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // If the Name is not empty, apply first‑character formatting.
        builder.Writeln("<<if [item.Name.Length > 0]>>");
        // First character in red color.
        builder.Writeln("<<textColor [\"Red\"]>><<[item.Name.Substring(0,1)]>><</textColor>><<[item.Name.Substring(1)]>>");
        builder.Writeln("<</if>>");

        // New line after each item.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(path);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
}
