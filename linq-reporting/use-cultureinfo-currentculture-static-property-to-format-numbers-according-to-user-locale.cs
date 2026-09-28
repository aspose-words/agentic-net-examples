using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Set the culture to French (France) to demonstrate locale‑specific number formatting.
        CultureInfo.CurrentCulture = new CultureInfo("fr-FR");

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Widget", Amount = 1234.56m },
                new Item { Name = "Gadget", Amount = 7890.12m },
                new Item { Name = "Doohickey", Amount = 345.67m }
            }
        };

        // Create the template document programmatically.
        var templatePath = "report_template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Sales Report");
        builder.Writeln($"Culture: {CultureInfo.CurrentCulture.Name}");
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.EndRow();

        // Table rows – format amount using the current culture.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Amount.ToString(\"N\")]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "report_output.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated with culture '{CultureInfo.CurrentCulture.Name}'.");
        Console.WriteLine($"Template: {Path.GetFullPath(templatePath)}");
        Console.WriteLine($"Output:   {Path.GetFullPath(outputPath)}");
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public decimal Amount { get; set; }
}
