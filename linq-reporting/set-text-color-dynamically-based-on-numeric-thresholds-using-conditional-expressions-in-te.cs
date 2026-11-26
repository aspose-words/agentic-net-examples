using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Value { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Dynamic Text Color Report");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Apply textColor based on numeric thresholds.
        // >80 => Red, >50 => Orange, otherwise Green.
        builder.Writeln(
            "<<textColor [item.Value > 80 ? \"Red\" : item.Value > 50 ? \"Orange\" : \"Green\"]>><<[item.Value]>> <</textColor>>");

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Value = 30 },
                new Item { Value = 55 },
                new Item { Value = 90 }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
