using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Value { get; set; }

    public Item(string name, int value)
    {
        Name = name;
        Value = value;
    }
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
        var model = new ReportModel();
        model.Items.Add(new Item("First", 10));
        model.Items.Add(new Item("Second", 20));
        model.Items.Add(new Item("Third", 30));
        model.Items.Add(new Item("Fourth", 40));

        // Create a template document programmatically.
        var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags that use ElementAt to fetch the third item (index 2).
        builder.Writeln("Third item name: <<[model.Items.ElementAt(2).Name]>>");
        builder.Writeln("Third item value: <<[model.Items.ElementAt(2).Value]>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        template.Save(outputPath);
    }
}
