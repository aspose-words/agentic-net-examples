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
        // Sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Create a template document.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Begin the foreach block.
        builder.Writeln("<<foreach [item in Items]>>");

        // Build the table inside the foreach block.
        builder.StartTable();

        // Header row – two cells merged horizontally.
        builder.InsertCell();
        builder.Writeln("<<cellMerge -horz>>Product Info");
        builder.InsertCell();
        builder.Writeln("<<cellMerge -horz>>Product Info");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template and generate the report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        reportDoc.Save(outputPath);
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
