using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words on .NET Core)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create template document with a table that merges cells horizontally
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Build a simple table
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Category");
        builder.InsertCell();
        builder.Writeln("Item");
        builder.EndRow();

        // Row with horizontally merged cells
        builder.InsertCell();
        builder.Writeln("<<cellMerge -horz>>Group A");
        builder.InsertCell();
        builder.Writeln("<<cellMerge -horz>>Group A");
        builder.EndRow();

        builder.EndTable();

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Sample data model (not used in this simple merge example)
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Category = "Group A", Name = "Item 1" },
                new Item { Category = "Group A", Name = "Item 2" }
            }
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        string outputPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(outputPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}
