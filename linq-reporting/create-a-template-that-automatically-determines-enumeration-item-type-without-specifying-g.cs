using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a foreach tag that iterates over a collection named Items.
        // The engine will infer the item type at runtime.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare sample data.
        // -------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // -------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}

// Public data model exposed to the template.
public class ReportModel
{
    // Strongly typed collection allows the engine to determine the item type automatically.
    public IEnumerable<Item> Items { get; set; } = new List<Item>();
}

// Item type used in the collection.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
