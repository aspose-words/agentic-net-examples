using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  BgColor = "LightYellow" },
                new Item { Name = "Banana", BgColor = "LightGreen" },
                new Item { Name = "Cherry", BgColor = "LightCoral" },
                new Item { Name = "Date",   BgColor = "LightGray" }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<backColor [item.BgColor]>>Item: <<[item.Name]>> <</backColor>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the template and the data model.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save("Report.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the foreach loop.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public string BgColor { get; set; } = string.Empty;
}
