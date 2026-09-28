using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create the template document programmatically.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Declare a variable for grand total.
        builder.Writeln("<<[var grandTotal = 0]>>");

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Calculate line total and add to grand total.
        builder.Writeln("<<[var lineTotal = item.Price * item.Quantity]>>");
        builder.Writeln("<<[grandTotal = grandTotal + lineTotal]>>");

        // Output item details.
        builder.Writeln(
            "Item: <<[item.Name]>>  Qty: <<[item.Quantity]>>  Price: <<[item.Price]>>  Total: <<[lineTotal]>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Output grand total after the loop.
        builder.Writeln("Grand Total: <<[grandTotal]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 3, Price = 0.5m },
                new Item { Name = "Banana", Quantity = 5, Price = 0.3m },
                new Item { Name = "Cherry", Quantity = 10, Price = 0.2m }
            }
        };

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine(success
            ? $"Report generated successfully: {Path.GetFullPath(outputPath)}"
            : "Report generation failed.");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}
