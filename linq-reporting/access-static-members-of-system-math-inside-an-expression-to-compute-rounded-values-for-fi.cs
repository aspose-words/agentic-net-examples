using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public decimal Amount { get; set; }

    // Computed property that returns the amount rounded to 2 decimal places.
    public decimal RoundedAmount => Math.Round(Amount, 2);
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create the template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Financial Report");
        builder.Writeln();
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Name]>>: $<<[item.RoundedAmount]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel();
        model.Items.Add(new Item { Name = "Product A", Amount = 123.4567m });
        model.Items.Add(new Item { Name = "Product B", Amount = 89.1234m });
        model.Items.Add(new Item { Name = "Service C", Amount = 45.6789m });

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
