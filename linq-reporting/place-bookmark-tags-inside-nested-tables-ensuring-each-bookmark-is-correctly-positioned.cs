using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and result documents
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "output.docx");

        // Create the template document with nested tables and bookmark tags
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin outer foreach over Sections
        builder.Writeln("<<foreach [section in Sections]>>");

        // Start outer table
        Table outerTable = builder.StartTable();

        // First cell: Section title
        builder.InsertCell();
        builder.Writeln("<<[section.Title]>>");

        // Second cell: will contain inner table with items
        builder.InsertCell();

        // Begin inner foreach over Items
        builder.Writeln("<<foreach [item in section.Items]>>");

        // Start inner table
        Table innerTable = builder.StartTable();

        // Cell with bookmark around item name
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");

        // End inner row and table
        builder.EndRow();
        builder.EndTable();

        // End inner foreach
        builder.Writeln("<</foreach>>");

        // End outer row and table
        builder.EndRow();
        builder.EndTable();

        // End outer foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Prepare sample data
        ReportModel model = new()
        {
            Sections = new()
            {
                new Section
                {
                    Title = "Section A",
                    Items = new()
                    {
                        new Item { Name = "Item A1", BookmarkName = "BM_A1" },
                        new Item { Name = "Item A2", BookmarkName = "BM_A2" }
                    }
                },
                new Section
                {
                    Title = "Section B",
                    Items = new()
                    {
                        new Item { Name = "Item B1", BookmarkName = "BM_B1" },
                        new Item { Name = "Item B2", BookmarkName = "BM_B2" },
                        new Item { Name = "Item B3", BookmarkName = "BM_B3" }
                    }
                }
            }
        };

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new();

        // Build the report
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        reportDoc.Save(resultPath);

        // Output result information
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {resultPath}");
    }
}

// Data model classes
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}
