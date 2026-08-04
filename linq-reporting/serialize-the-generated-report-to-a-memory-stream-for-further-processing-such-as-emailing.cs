using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Simple title.
        builder.Writeln("Order Report");
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln();

        // LINQ Reporting foreach block to list items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>> - Qty: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the data model.
        ReportingEngine engine = new ReportingEngine();
        OrderReport model = CreateSampleModel();
        engine.BuildReport(template, model, "model");

        // Serialize the generated report to a memory stream (e.g., for emailing).
        using (MemoryStream reportStream = new MemoryStream())
        {
            template.Save(reportStream, SaveFormat.Docx);
            // The stream now contains the DOCX bytes.
            Console.WriteLine($"Report generated. Stream length: {reportStream.Length} bytes");
            // Reset position if the stream will be read later.
            reportStream.Position = 0;
            // Further processing such as attaching to an email would use 'reportStream'.
        }
    }

    // Creates a sample data model for the report.
    private static OrderReport CreateSampleModel()
    {
        return new OrderReport
        {
            CustomerName = "John Doe",
            Items = new()
            {
                new Item { Name = "Apple", Quantity = 3 },
                new Item { Name = "Banana", Quantity = 5 },
                new Item { Name = "Orange", Quantity = 2 }
            }
        };
    }

    // Root data model referenced in the template as 'model'.
    public class OrderReport
    {
        public string CustomerName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    // Item model used inside the foreach loop.
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }
}
