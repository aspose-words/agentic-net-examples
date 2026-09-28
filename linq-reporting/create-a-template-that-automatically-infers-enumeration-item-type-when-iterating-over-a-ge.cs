using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
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
        model.Items.Add(new Item { Index = 1, Name = "Alice" });
        model.Items.Add(new Item { Index = 2, Name = "Bob" });
        model.Items.Add(new Item { Index = 3, Name = "Charlie" });

        // Create the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("LINQ Reporting Example");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
