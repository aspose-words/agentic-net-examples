using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "ReportOutput.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Sample Report ===");
        // Start a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        // Always show the item name.
        builder.Writeln("Item: <<[item.Name]>>");
        // Show the value only when it exceeds the threshold (100).
        builder.Writeln("<<if [item.Value > 100]>>Value exceeds threshold: <<[item.Value]>> <</if>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new()
            {
                new Item { Name = "Alpha",   Value = 75 },
                new Item { Name = "Beta",    Value = 120 },
                new Item { Name = "Gamma",   Value = 95 },
                new Item { Name = "Delta",   Value = 180 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The root object is 'model', and the template references its Items collection directly.
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection of items to be displayed in the report.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    // Name of the item.
    public string Name { get; set; } = string.Empty;

    // Numeric value associated with the item.
    public double Value { get; set; }
}
