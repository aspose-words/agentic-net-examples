using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Product Report");
        builder.Writeln();

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Always show the item name.
        builder.Writeln("Item: <<[item.Name]>>");

        // Conditional section: show the value only if it exceeds the threshold (e.g., 50).
        builder.Writeln("<<if [item.Value > 50]>>");
        builder.Writeln("  Value exceeds threshold: <<[item.Value]>>");
        builder.Writeln("<</if>>");

        // End of foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Widget A", Value = 30 },
                new Item { Name = "Widget B", Value = 75 },
                new Item { Name = "Widget C", Value = 120 },
                new Item { Name = "Widget D", Value = 45 }
            }
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Remove empty paragraphs that may result from false conditions.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report using the root object name "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection of items to be iterated in the template.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    // Name of the product.
    public string Name { get; set; } = string.Empty;

    // Numeric value used for the conditional check.
    public int Value { get; set; }
}
