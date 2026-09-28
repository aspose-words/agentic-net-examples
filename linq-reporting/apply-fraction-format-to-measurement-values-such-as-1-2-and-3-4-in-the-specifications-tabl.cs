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
        // Ensure the output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Specifications");
        builder.Writeln();

        // Begin the foreach block for items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table for each item (header + data row).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Measurement");
        builder.EndRow();

        // Data row (repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Fraction]>>");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new() { Index = 1, Name = "Widget A", Measurement = 0.5 },
                new() { Index = 2, Name = "Widget B", Measurement = 0.75 },
                new() { Index = 3, Name = "Widget C", Measurement = 0.25 },
                new() { Index = 4, Name = "Widget D", Measurement = 1.0 }
            }
        };

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "output.docx");
        doc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model with fraction formatting.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public double Measurement { get; set; }

    // Returns a fraction string for known measurement values.
    public string Fraction
    {
        get
        {
            return Measurement switch
            {
                double m when Math.Abs(m - 0.5) < 0.0001 => "1/2",
                double m when Math.Abs(m - 0.75) < 0.0001 => "3/4",
                double m when Math.Abs(m - 0.25) < 0.0001 => "1/4",
                double m when Math.Abs(m - 1.0) < 0.0001 => "1",
                _ => Measurement.ToString()
            };
        }
    }
}
