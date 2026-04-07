using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Build the template:
        //   <<foreach [item in Items]>>
        //       <<if [item.Stock > 0 && item.Price < 100]>>
        //           <<[item.Name]>> - Stock: <<[item.Stock]>> - Price: <<[item.Price]>>
        //       <</if>>
        //   <</foreach>>
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("<<[item.Name]>> - Stock: <<[item.Stock]>> - Price: <<[item.Price]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple",  Stock = 50, Price = 0.75 },
                new Item { Name = "Banana", Stock = 0,  Price = 0.30 },
                new Item { Name = "Cherry", Stock = 20, Price = 1.20 },
                new Item { Name = "Date",   Stock = 15, Price = 2.50 }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        template.Save("Report.docx");
    }
}

// Wrapper class that matches the root object name used in BuildReport ("model").
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple data entity used in the report.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Stock { get; set; }
    public double Price { get; set; }
}
