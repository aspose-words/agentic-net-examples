using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with a table that repeats for each item.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach block that iterates over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start the table (header will be written once before the loop).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Link");
        builder.EndRow();

        // Data row – will be repeated for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        // Link tag: first expression is the URL, second is the display text.
        builder.Writeln("<<link [item.Url] [item.Name]>>");
        builder.EndRow();

        // Finish the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 2. Load the template and prepare the data source for the report.
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Name = "Aspose", Url = "https://www.aspose.com" },
                new Item { Name = "GitHub", Url = "https://github.com" }
            }
        };

        // -------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None // default options
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes – must be public with public properties.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
}
