using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words on some platforms)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var order = new Order
        {
            CustomerName = "John Doe",
            Total = 1234.56m,
            Items = new List<Item>
            {
                new Item { Name = "Widget", Price = 123.45m },
                new Item { Name = "Gadget", Price = 250.00m }
            }
        };

        // Serialize data to JSON file
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "order.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(order, Formatting.Indented));

        // Create a template document with LINQ Reporting tags
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Invoice");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total Amount: <<[order.Total.ToString(\"C\")]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Table for items
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Price.ToString(\"C\")]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection)
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Build the report using the JSON data source
        // Load JSON data back (demonstrates JsonDataSource usage)
        string jsonContent = File.ReadAllText(jsonPath);
        var data = JsonConvert.DeserializeObject<Order>(jsonContent) ?? new Order();

        var engine = new ReportingEngine();
        engine.BuildReport(template, data, "order");

        // Save the generated report
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        template.Save(reportPath);
    }
}

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal Total { get; set; }
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
