using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Program
{
    public static void Main()
    {
        // Sample data model.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Value = 10 },
                new Item { Name = "Banana", Value = 20 },
                new Item { Name = "Cherry", Value = 30 }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder(); // Creates a new blank document.

        // Open a foreach block before the table.
        builder.Writeln("<<foreach [item in Items]>>");

        // Build the table that will be repeated for each item.
        Table table = builder.StartTable();

        // Header row – two horizontally merged cells.
        builder.InsertCell();
        builder.Write("<<cellMerge>>Category");
        builder.InsertCell();
        builder.Write("<<cellMerge>>Category");
        builder.EndRow();

        // Data row – tags will be replaced for each item.
        builder.InsertCell();
        builder.Write("<<[item.Name]>>");
        builder.InsertCell();
        builder.Write("<<[item.Value]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        builder.Document.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Populate the template with the data model (root name "model").
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
}
