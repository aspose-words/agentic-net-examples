using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with a table that contains LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder();

        // Open a foreach block before the table so that the whole table is repeated.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start the table that will be repeated for each item.
        Table table = builder.StartTable();

        // First row – cells that will be merged horizontally.
        builder.InsertCell();
        builder.Writeln("<<cellMerge -horz>>"); // Marks this cell for horizontal merge.
        builder.InsertCell();
        builder.Writeln("<<cellMerge -horz>>"); // Marks the adjacent cell for horizontal merge.
        builder.EndRow();

        // Second row – data row that will be repeated for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Value]>>");
        builder.EndRow();

        // End the table and close the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        builder.Document.Save(templatePath);

        // 2. Load the template for reporting.
        var doc = new Document(templatePath);

        // 3. Prepare the data source.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Item A", Value = "10" },
                new Item { Name = "Item B", Value = "20" },
                new Item { Name = "Item C", Value = "30" }
            }
        };

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None // No special options required.
        };
        bool success = engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate success.
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}

// Data model classes must be public with public properties.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
