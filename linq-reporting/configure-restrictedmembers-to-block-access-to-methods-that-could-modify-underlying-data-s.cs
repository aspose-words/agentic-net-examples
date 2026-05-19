using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple" },
                new Item { Name = "Banana" }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Attempt to modify the data source inside the template (this will be blocked).
        builder.Writeln("Attempt to add an item (should be blocked):");
        builder.Writeln("<<[Items.Add(new Item { Name = \"Cherry\" })]>>");

        // Normal foreach loop to display items.
        builder.Writeln("Items list:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Restrict access to List<Item> so that its mutating members are unavailable.
        // -----------------------------------------------------------------
        ReportingEngine.SetRestrictedTypes(typeof(List<Item>));

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            // Inline error messages will show that the blocked call could not be executed.
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "report.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive input required).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// ---------------------------------------------------------------------
// Data model classes (public with public properties, non‑nullable).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
}
