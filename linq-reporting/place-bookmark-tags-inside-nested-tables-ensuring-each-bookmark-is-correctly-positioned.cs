using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for Table type

public class Program
{
    public static void Main()
    {
        // Sample data model.
        ReportModel model = new()
        {
            Sections = new List<Section>
            {
                new()
                {
                    Title = "First Section",
                    Items = new List<Item>
                    {
                        new() { Name = "Item 1A", BookmarkName = "BM_1A" },
                        new() { Name = "Item 1B", BookmarkName = "BM_1B" }
                    }
                },
                new()
                {
                    Title = "Second Section",
                    Items = new List<Item>
                    {
                        new() { Name = "Item 2A", BookmarkName = "BM_2A" },
                        new() { Name = "Item 2B", BookmarkName = "BM_2B" },
                        new() { Name = "Item 2C", BookmarkName = "BM_2C" }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin outer foreach over sections.
        builder.Writeln("<<foreach [section in Sections]>>");
        builder.Writeln("Section: <<[section.Title]>>");

        // Outer table – one cell per section that will contain the inner table.
        Table outerTable = builder.StartTable();
        builder.InsertCell(); // start outer cell

        // Begin inner foreach over items of the current section.
        builder.Writeln("<<foreach [item in section.Items]>>");

        // Inner table – each item becomes a row with a bookmark.
        Table innerTable = builder.StartTable();

        // Row for the current item.
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");
        builder.EndRow();

        builder.EndTable(); // end inner table for this item
        builder.Writeln("<</foreach>>"); // end inner foreach

        // End outer cell/row and outer table.
        builder.EndRow();
        builder.EndTable();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to a temporary file.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
