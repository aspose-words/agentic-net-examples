using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with a LINQ Reporting foreach tag.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // The template iterates over the pre‑transformed collection.
        builder.Writeln("<<foreach [s in Transformed]>>");
        builder.Writeln("<<[s]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 2. Load the template and prepare the data model for the report.
        // ---------------------------------------------------------------
        var doc = new Document(templatePath);

        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Id = 1, Name = "Apple" },
                new Item { Id = 2, Name = "Banana" },
                new Item { Id = 3, Name = "Cherry" }
            }
        };

        // -------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model exposed to the template.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection of source items.
    public List<Item> Items { get; set; } = new();

    // Custom delegate used for formatting each item.
    public Func<Item, string> Transform => item => $"Id:{item.Id} Name:{item.Name}";

    // Helper property that applies the delegate to each item.
    // The template iterates over this property.
    public IEnumerable<string> Transformed => Items.Select(Transform);
}

// Simple data item class.
public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}
