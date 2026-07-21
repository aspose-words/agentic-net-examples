using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alice", Score = 92 },
                new Item { Name = "Bob",   Score = 67 },
                new Item { Name = "Carol", Score = 45 }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Dynamic Text Color Report");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(
            "<<textColor [item.Score >= 80 ? \"Red\" : item.Score >= 50 ? \"Orange\" : \"Green\"]>>" +
            "Name: <<[item.Name]>>  Score: <<[item.Score]>>" +
            "<</textColor>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Score { get; set; }
}
