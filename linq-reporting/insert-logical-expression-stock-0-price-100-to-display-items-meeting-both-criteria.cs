using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Stock { get; set; }
    public decimal Price { get; set; }
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
                new Item { Name = "Apple", Stock = 10, Price = 0.99m },
                new Item { Name = "Banana", Stock = 0, Price = 0.59m },
                new Item { Name = "Cherry", Stock = 5, Price = 120.00m },
                new Item { Name = "Date", Stock = 3, Price = 45.50m }
            }
        };

        // Create the template document.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Items with Stock > 0 and Price < 100:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("Name: <<[item.Name]>>, Stock: <<[item.Stock]>>, Price: <<[item.Price]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save and reload the template to simulate real scenario.
        const string templatePath = "Template.docx";
        template.Save(templatePath);
        var loadedTemplate = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}
