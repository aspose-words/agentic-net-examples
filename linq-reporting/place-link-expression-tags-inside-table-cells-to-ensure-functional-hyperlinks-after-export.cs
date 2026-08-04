using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Url = "https://example.com/1", Text = "Example 1" },
                new Item { Url = "https://example.com/2", Text = "Example 2" },
                new Item { Url = "https://example.com/3", Text = "Example 3" }
            }
        };

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with two columns: a header and a hyperlink column.
        var table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Link");
        builder.EndRow();

        // Data row (repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<[item.Text]>>");
        builder.InsertCell();
        // Place the link tag inside the cell.
        builder.Writeln("<<link [item.Url] [item.Text]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the data source.
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
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
    public string Url { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
}
