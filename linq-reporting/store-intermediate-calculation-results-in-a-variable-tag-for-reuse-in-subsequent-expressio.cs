using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder();

        // Begin iterating over the collection of items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Output item details.
        builder.Writeln("Item: <<[item.Name]>> - Price: <<[item.Price]>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Output the accumulated total using a LINQ expression.
        builder.Writeln("Total Price: <<[Items.Sum(item => item.Price)]>>");

        // Save the template.
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare the data source.
        var order = new Order
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 1.20 },
                new Item { Name = "Banana", Price = 0.80 },
                new Item { Name = "Orange", Price = 1.50 }
            }
        };

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine { Options = ReportBuildOptions.None };
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Public data model classes.
public class Order
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
