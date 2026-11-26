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
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create the template document with a table and LINQ Reporting tags
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach loop over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Table that will be repeated for each item
        Table table = builder.StartTable();

        // Header row (appears for each item – acceptable for this simple example)
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Data row with conditional background color
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsImportant]>>" +
            "<<backColor [\"Yellow\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.IsImportant == false]>>" +
            "<<[item.Name]>>" +
            "<</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsImportant]>>" +
            "<<backColor [\"Yellow\"]>><<[item.Value]>> <</backColor>><</if>>" +
            "<<if [item.IsImportant == false]>>" +
            "<<[item.Value]>>" +
            "<</if>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Alpha",   Value = 100, IsImportant = true },
                new Item { Name = "Beta",    Value = 200, IsImportant = false },
                new Item { Name = "Gamma",   Value = 300, IsImportant = true },
                new Item { Name = "Delta",   Value = 400, IsImportant = false },
                new Item { Name = "Epsilon", Value = 500, IsImportant = true }
            }
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document
        string outputPath = Path.Combine(outputDir, "output.docx");
        reportDoc.Save(outputPath);
    }
}

// Root data model
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item data class
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
    public bool IsImportant { get; set; }
}
