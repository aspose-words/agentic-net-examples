using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class OrderItem
{
    public string Description { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}

public class OrderReport
{
    public List<OrderItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new(templatePath);

        // Prepare sample data.
        OrderReport model = new()
        {
            Items = new()
            {
                new OrderItem { Description = "Widget A", Quantity = 3, UnitPrice = 9.99m },
                new OrderItem { Description = "Widget B", Quantity = 5, UnitPrice = 14.50m },
                new OrderItem { Description = "Widget C", Quantity = 2, UnitPrice = 7.25m }
            }
        };

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Title
        builder.Writeln("Order Items Report");
        builder.Writeln();

        // Begin foreach band.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start table inside foreach.
        Table table = builder.StartTable();

        // Header row (only once, before data rows).
        builder.InsertCell();
        builder.Writeln("Description");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Unit Price");
        builder.InsertCell();
        builder.Writeln("Line Total");
        builder.EndRow();

        // Data row inside foreach.
        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.UnitPrice]>>");
        builder.InsertCell();
        // Calculate line total using expression tag.
        builder.Writeln("<<[item.Quantity * item.UnitPrice]>>");
        builder.EndRow();

        // End table.
        builder.EndTable();

        // End foreach band.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(path);
    }
}
