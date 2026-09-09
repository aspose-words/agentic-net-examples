using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
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
        // Sample data
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Stock = 10, Price = 50 },
                new Item { Name = "Banana", Stock = 0,  Price = 30 },
                new Item { Name = "Cherry", Stock = 5,  Price = 120 },
                new Item { Name = "Date",   Stock = 3,  Price = 80 }
            }
        };

        // Create template document
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Items with Stock > 0 and Price < 100:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("- <<[item.Name]>> : Stock = <<[item.Stock]>>, Price = $<<[item.Price]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load template and generate report
        var template = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save final report
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
