using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public int Value { get; set; }
    public string Description { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create template document.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Dynamic Text Color Report");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Write a line with value colored based on thresholds.
        // Red > 100, Orange > 50, otherwise Green.
        builder.Writeln(
            "Value: " +
            "<<textColor [item.Value > 100 ? \"Red\" : (item.Value > 50 ? \"Orange\" : \"Green\")]>>" +
            "<<[item.Value]>>" +
            "<</textColor>>");

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Value = 30, Description = "Low" },
                new Item { Value = 75, Description = "Medium" },
                new Item { Value = 120, Description = "High" }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "output.docx";
        template.Save(outputPath);
    }
}
