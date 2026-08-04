using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Program
{
    public static void Main()
    {
        // Set the current culture to demonstrate locale‑specific number formatting.
        CultureInfo.CurrentCulture = new CultureInfo("fr-FR");

        // Create a simple template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin the foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");

        // Build the table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Build the data rows.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        // The model provides a locale‑aware formatted string.
        builder.Writeln("<<[item.FormattedPrice]>>");
        builder.EndRow();
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for completeness).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 1.23m },
                new Item { Name = "Banana", Price = 0.99m },
                new Item { Name = "Cherry", Price = 2.50m }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item with a locale‑aware formatted price.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }

    // Returns the price formatted according to the current culture.
    public string FormattedPrice => Price.ToString("C", CultureInfo.CurrentCulture);
}
