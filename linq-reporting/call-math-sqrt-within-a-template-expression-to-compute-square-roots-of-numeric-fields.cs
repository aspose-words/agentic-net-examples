using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public double Value { get; set; }

    // Computed property that returns the square root of Value.
    public double Sqrt => Math.Sqrt(Value);
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Value = 4 },
                new Item { Value = 9 },
                new Item { Value = 16 },
                new Item { Value = 25 }
            }
        };

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Square Root Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Original: <<[item.Value]>>   Sqrt: <<[item.Sqrt]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        template.Save(outputPath);
    }
}
