using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class Model
{
    // Strongly‑typed collection so the LINQ Reporting engine can infer the item type.
    public IEnumerable<Item> Items { get; set; } = new List<Item>();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var items = new List<Item>
        {
            new() { Index = 1, Name = "Alice" },
            new() { Index = 2, Name = "Bob" },
            new() { Index = 3, Name = "Charlie" }
        };

        var model = new Model { Items = items };

        // Create the template document.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Index: <<[item.Index]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
