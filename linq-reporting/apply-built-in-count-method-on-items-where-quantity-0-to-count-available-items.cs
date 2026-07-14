using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
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
        model.Items.Add(new Item { Name = "Apple", Quantity = 5 });
        model.Items.Add(new Item { Name = "Banana", Quantity = 0 });
        model.Items.Add(new Item { Name = "Orange", Quantity = 3 });

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Available items count: <<[model.Items.Count(i => i.Quantity > 0)]>>");
        // Save the template to a local file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
