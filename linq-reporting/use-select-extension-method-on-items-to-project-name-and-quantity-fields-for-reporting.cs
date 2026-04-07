using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class ReportItem
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a heading.
        builder.Writeln("Items Report");
        builder.Writeln();

        // LINQ Reporting foreach tag iterating over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>  |  Quantity: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare data source.
        // -----------------------------------------------------------------
        var items = new List<Item>
        {
            new() { Name = "Apple",  Quantity = 10, Price = 0.5m },
            new() { Name = "Banana", Quantity = 20, Price = 0.3m },
            new() { Name = "Orange", Quantity = 15, Price = 0.4m }
        };

        // Use LINQ Select to project only Name and Quantity fields.
        var projected = items
            .Select(i => new ReportItem { Name = i.Name, Quantity = i.Quantity })
            .ToList();

        var model = new ReportModel { Items = projected };

        // -----------------------------------------------------------------
        // 4. Build the report using Aspose.Words LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}
