using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        Order order = new()
        {
            OrderId = 12345,
            CustomerName = "John Doe",
            Items = new()
            {
                new OrderItem { ProductName = "Widget", Quantity = 3 },
                new OrderItem { ProductName = "Gadget", Quantity = 5 }
            }
        };

        // Wrap data for reporting
        ReportModel model = new()
        {
            Order = order,
            OrderJson = JsonConvert.SerializeObject(order, Formatting.Indented)
        };

        // Create template document programmatically
        Document template = new();
        DocumentBuilder builder = new(template);
        builder.Writeln("Debug Order JSON:");
        builder.Writeln("<<[model.OrderJson]>>");

        // Save template to disk
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load template for report generation
        Document doc = new(templatePath);

        // Build report using LINQ Reporting engine
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Wrapper model for the report
public class ReportModel
{
    public Order Order { get; set; } = null!;
    public string OrderJson { get; set; } = string.Empty;
}

// Sample order data model
public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public List<OrderItem> Items { get; set; } = new();
}

// Sample order item model
public class OrderItem
{
    public string ProductName { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
