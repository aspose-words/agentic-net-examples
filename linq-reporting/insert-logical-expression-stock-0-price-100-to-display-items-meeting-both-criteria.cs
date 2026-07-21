using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Build a LINQ Reporting template that iterates over Items and applies a logical filter.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("Item: <<[item.Name]>> | Stock: <<[item.Stock]>> | Price: <<[item.Price]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk before building the report.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the saved template.
        Document loadedTemplate = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Stock = 10, Price =  50.0 },
                new Item { Name = "Banana", Stock = 0,  Price =  30.0 },
                new Item { Name = "Cherry", Stock = 5,  Price = 150.0 },
                new Item { Name = "Date",   Stock = 3,  Price =  80.0 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}

// Root data model referenced in the template as "model".
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item class used inside the foreach loop.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Stock { get; set; }
    public double Price { get; set; }
}
