using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample hierarchical data.
        ReportModel model = new ReportModel
        {
            Categories = new List<Category>
            {
                new Category
                {
                    Name = "Fruits",
                    Items = new List<Item>
                    {
                        new Item { Index = 1, Name = "Apple" },
                        new Item { Index = 2, Name = "Banana" }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new List<Item>
                    {
                        new Item { Index = 1, Name = "Carrot" },
                        new Item { Index = 2, Name = "Tomato" }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Outer foreach over categories.
        builder.Writeln("<<foreach [category in Categories]>>");
        builder.Writeln("Category: <<[category.Name]>>");

        // Inner foreach over items of the current category.
        builder.Writeln("<<foreach [item in category.Items]>>");
        builder.Writeln("- Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>"); // End inner foreach.

        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report using the data model.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options needed.

        // Build the report. The root object name must match the name used in the template tags.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Optionally, you could check the success flag if InlineErrorMessages were enabled.
        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public with public properties, no nullable warnings).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
