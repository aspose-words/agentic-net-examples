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
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Stock = 10, Price =  50 },
                new Item { Name = "Banana", Stock = 0,  Price =  30 },
                new Item { Name = "Cherry", Stock = 5,  Price = 150 },
                new Item { Name = "Date",   Stock = 3,  Price =  80 }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("Item: <<[item.Name]>> | Stock: <<[item.Stock]>> | Price: <<[item.Price]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, demonstrates load/save lifecycle).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template (could reuse the same document, shown for completeness).
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
