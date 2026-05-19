using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Stock { get; set; }
    public double Price { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Header
        builder.Writeln("Items with Stock > 0 and Price < 100:");
        builder.Writeln();

        // Begin foreach over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Conditional block: display only items meeting both criteria.
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("- <<[item.Name]>> (Stock: <<[item.Stock]>>, Price: $<<[item.Price]>>)");
        builder.Writeln("<</if>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // 2. Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Name = "Apple", Stock = 20, Price = 45.5 },
                new Item { Name = "Banana", Stock = 0, Price = 30.0 },
                new Item { Name = "Cherry", Stock = 15, Price = 120.0 },
                new Item { Name = "Date", Stock = 5, Price = 80.0 }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        reportDoc.Save("Report.docx");
    }
}
