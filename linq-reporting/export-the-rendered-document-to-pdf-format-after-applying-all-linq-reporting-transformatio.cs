using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        Order order = new()
        {
            CustomerName = "John Doe",
            Items = new()
            {
                new Item { Name = "Apple", Quantity = 3 },
                new Item { Name = "Banana", Quantity = 5 },
                new Item { Name = "Orange", Quantity = 2 }
            }
        };

        // Create a template document programmatically.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, order, "order");

        // Export the rendered document to PDF.
        string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Write static text and LINQ Reporting tags.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();
        builder.Writeln("Items:");

        // Begin foreach block for items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start table inside the foreach block.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data row (repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();

        // End table.
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(path);
    }
}

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
