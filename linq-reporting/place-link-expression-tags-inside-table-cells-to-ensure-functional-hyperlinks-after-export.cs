using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table class

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "ReportWithLinks.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table for each item.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Link");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        // Insert the item's name.
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        // Insert a functional hyperlink using the link tag.
        builder.Writeln("<<link [item.Url] [item.LinkText]>>");
        builder.EndRow();

        // Close the table.
        builder.EndTable();

        // End the foreach block.
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
                new Item
                {
                    Name = "Aspose",
                    Url = "https://www.aspose.com",
                    LinkText = "Visit Aspose"
                },
                new Item
                {
                    Name = "GitHub",
                    Url = "https://github.com",
                    LinkText = "Open GitHub"
                },
                new Item
                {
                    Name = "Stack Overflow",
                    Url = "https://stackoverflow.com",
                    LinkText = "Go to Stack Overflow"
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties are initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
    public string LinkText { get; set; } = string.Empty;
}
