using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public double Number { get; set; } = 0;
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

        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Number: <<[item.Number]>>");
        builder.Writeln("Square root: <<[Math.Sqrt(item.Number)]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template.
        var loadedDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Number = 1 },
                new Item { Number = 4 },
                new Item { Number = 9 },
                new Item { Number = 16 }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Math));
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        loadedDoc.Save("Report.docx");
    }
}
