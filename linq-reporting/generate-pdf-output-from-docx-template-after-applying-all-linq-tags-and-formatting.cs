using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the DOCX template with LINQ tags.
        const string templateFile = "Template.docx";
        CreateTemplate(templateFile);

        // Load the template.
        var doc = new Document(templateFile);

        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 3 },
                new Item { Name = "Banana", Quantity = 5 },
                new Item { Name = "Cherry", Quantity = 7 }
            }
        };

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine { Options = ReportBuildOptions.RemoveEmptyParagraphs };
        engine.BuildReport(doc, order, "order");

        // Save the populated document as PDF.
        doc.Save("Report.pdf", SaveFormat.Pdf);
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ tags into the template.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>> : <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Public data model classes.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
