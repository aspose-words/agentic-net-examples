using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var reportData = new ReportModel
        {
            Title = "Order Summary",
            Order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Apple" },
                    new Item { Index = 2, Name = "Banana" },
                    new Item { Index = 3, Name = "Cherry" }
                }
            }
        };

        // Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("Customer: <<[model.Order.CustomerName]>>");
        builder.Writeln();
        builder.Writeln("<<foreach [item in model.Order.Items]>>");

        // Table header.
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Product");
        builder.EndRow();

        // Table row for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, reportData, "model");

        // Serialize the generated report to a memory stream.
        using var reportStream = new MemoryStream();
        templateDoc.Save(reportStream, SaveFormat.Docx);
        reportStream.Position = 0; // Reset for further processing.

        // For demonstration, save the stream to a file.
        const string outputPath = "GeneratedReport.docx";
        File.WriteAllBytes(outputPath, reportStream.ToArray());

        // Optionally, display the size of the generated report.
        Console.WriteLine($"Report generated and saved to '{outputPath}'. Size: {reportStream.Length} bytes.");
    }
}

// Root model class referenced in the template as "model".
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public Order Order { get; set; } = new();
}

// Order class containing customer information and a collection of items.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Item class used inside the foreach loop.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
