using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { DurationString = "02:30:00" }, // 2 hours 30 minutes
                new Item { DurationString = "00:45:15" }, // 45 minutes 15 seconds
                new Item { DurationString = "01:15:30" }  // 1 hour 15 minutes 30 seconds
            }
        };

        // Create a template document programmatically.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Write a heading.
        builder.Writeln("Duration Report");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with two columns: original string and parsed TimeSpan.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Original String");
        builder.InsertCell();
        builder.Writeln("Parsed TimeSpan");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.DurationString]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Duration]>>");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to a temporary file.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        doc.Save(templatePath);

        // Load the template for reporting.
        Document templateDoc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        templateDoc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model with duration parsing.
public class Item
{
    // Original duration string from data source.
    public string DurationString { get; set; } = string.Empty;

    // Parsed TimeSpan using TimeSpan.Parse.
    public TimeSpan Duration => TimeSpan.Parse(DurationString);
}
