using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product List:");
        builder.Writeln("<<foreach [item in Items]>>");

        // Build a simple two‑column table inside the foreach block.
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare sample data.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Demonstrate a foreach loop with implicit type inference.
        foreach (var item in model.Items)
        {
            Console.WriteLine($"{item.Index}: {item.Name}");
        }

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);
    }
}
