using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Table row for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to a file.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Serialize the generated report to a memory stream.
        using MemoryStream reportStream = new();
        reportDoc.Save(reportStream, SaveFormat.Docx);
        reportStream.Position = 0;

        // Example usage of the memory stream (e.g., write length to console).
        Console.WriteLine($"Report generated. Stream length: {reportStream.Length} bytes.");

        // Optionally, save the report to a file for verification.
        File.WriteAllBytes("Report.docx", reportStream.ToArray());
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
