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
            Sections = new List<Section>
            {
                new Section
                {
                    Name = "Section A",
                    Items = new List<Item>
                    {
                        new Item { BookmarkName = "BM_A1", Title = "Item A1 Title" },
                        new Item { BookmarkName = "BM_A2", Title = "Item A2 Title" }
                    }
                },
                new Section
                {
                    Name = "Section B",
                    Items = new List<Item>
                    {
                        new Item { BookmarkName = "BM_B1", Title = "Item B1 Title" },
                        new Item { BookmarkName = "BM_B2", Title = "Item B2 Title" }
                    }
                }
            }
        };

        // Create the template document programmatically.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin outer foreach over Sections.
        builder.Writeln("<<foreach [section in Sections]>>");

        // Start outer table.
        var outerTable = builder.StartTable();

        // First cell: Section name.
        builder.InsertCell();
        builder.Writeln("<<[section.Name]>>");

        // Second cell: inner table with items.
        builder.InsertCell();

        // Begin inner foreach over Items.
        builder.Writeln("<<foreach [item in section.Items]>>");

        // Start inner table.
        var innerTable = builder.StartTable();

        // Cell for bookmark content.
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");

        // End inner table row.
        builder.EndRow();

        // End inner table.
        builder.EndTable();

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // End outer table row.
        builder.EndRow();

        // End outer table.
        builder.EndTable();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string BookmarkName { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
}
