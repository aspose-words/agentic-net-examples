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
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  ColorName = "Red" },
                new Item { Name = "Banana", ColorName = "#FFD700" }, // gold
                new Item { Name = "Grape",  ColorName = "Purple" }
            }
        };

        // Create the template document programmatically.
        var templatePath = Path.Combine("Output", "Template.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(templatePath)!);
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report with dynamic text colors:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<textColor [item.ColorName]>>Item: <<[item.Name]>> <</textColor>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template (optional, shown for completeness).
        var loadedDoc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(loadedDoc, model, "model");

        // Save the final document to HTML while preserving the color tags.
        var htmlPath = Path.Combine("Output", "Report.html");
        var saveOptions = new HtmlSaveOptions(); // No ExportColorNames property needed.
        loadedDoc.Save(htmlPath, saveOptions);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item used in the foreach loop.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public string ColorName { get; set; } = string.Empty;
}
