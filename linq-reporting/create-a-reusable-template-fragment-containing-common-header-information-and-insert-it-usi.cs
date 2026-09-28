using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create reusable header fragment.
        var headerDoc = new Document();
        var headerBuilder = new DocumentBuilder(headerDoc);
        headerBuilder.Writeln("<<[model.Title]>>");
        headerBuilder.Writeln("Report Date: <<[model.ReportDate]>>");

        // Create main template and insert the header fragment.
        var mainDoc = new Document();
        var builder = new DocumentBuilder(mainDoc);
        builder.InsertDocument(headerDoc, ImportFormatMode.KeepSourceFormatting);
        builder.Writeln(); // Add a blank line after the header.

        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.EndRow();

        // Table row (repeated for each item).
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template (optional, just for demonstration).
        mainDoc.Save("MainTemplate.docx");

        // Sample data model.
        var model = new ReportModel
        {
            Title = "Sales Report",
            ReportDate = DateTime.Now.ToString("yyyy-MM-dd"),
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 10 },
                new Item { Name = "Banana", Quantity = 20 },
                new Item { Name = "Cherry", Quantity = 15 }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(mainDoc, model, "model");

        // Save the generated report.
        mainDoc.Save("ReportOutput.docx");
        Console.WriteLine("Report generated: ReportOutput.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string ReportDate { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
