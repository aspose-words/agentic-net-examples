using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; } = 0;
    public string CustomerName { get; set; } = "";
    public string Status { get; set; } = "";
    public decimal Amount { get; set; } = 0m;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        List<Order> orders = new()
        {
            new Order { Id = 1, CustomerName = "Alice",   Status = "Pending",   Amount = 120.50m },
            new Order { Id = 2, CustomerName = "Bob",     Status = "Shipped",   Amount =  85.00m },
            new Order { Id = 3, CustomerName = "Charlie", Status = "Pending",   Amount =  45.75m },
            new Order { Id = 4, CustomerName = "Diana",   Status = "Delivered", Amount = 210.00m }
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Orders Report (Pending Only)");
        // Use Where extension method inside the foreach tag to filter pending orders
        builder.Writeln("<<foreach [order in orders.Where(o => o.Status == \"Pending\")]>>");
        builder.Writeln("Order ID: <<[order.Id]>>, Customer: <<[order.CustomerName]>>, Amount: $<<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Pass the orders list as the data source named "orders"
        engine.BuildReport(report, orders, "orders");

        // Save the generated report
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
