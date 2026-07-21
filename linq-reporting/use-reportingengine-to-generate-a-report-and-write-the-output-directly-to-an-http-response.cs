using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    // Simple data model for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "Sample Item";
        public decimal Price { get; set; }
    }

    public static void Main()
    {
        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a title and a placeholder for the customer name.
        builder.Writeln("Order Report");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin a foreach loop over the order items.
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // 2. Prepare sample data.
        Order order = new Order
        {
            CustomerName = "Alice Smith",
            Items = new List<Item>
            {
                new Item { Name = "Laptop", Price = 1299.99m },
                new Item { Name = "Mouse", Price = 25.50m },
                new Item { Name = "Keyboard", Price = 45.00m }
            }
        };

        // 3. Build the report using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, order, "order");

        // 4. Save the generated document to a memory stream and then to a file.
        using (MemoryStream stream = new MemoryStream())
        {
            // Save directly to the stream in DOCX format.
            template.Save(stream, SaveFormat.Docx);

            // Write the stream contents to a physical file for demonstration.
            stream.Position = 0;
            using (FileStream file = File.Create("GeneratedReport.docx"))
            {
                stream.CopyTo(file);
            }
        }

        Console.WriteLine("Report generated and saved as 'GeneratedReport.docx'.");
    }
}
