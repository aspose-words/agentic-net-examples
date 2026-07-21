using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int OrderId { get; set; }
    public int CustomerId { get; set; }
    public string Product { get; set; } = "";
    public double Price { get; set; }
}

public class CustomerGroup
{
    public int CustomerId { get; set; }
    public List<Order> Orders { get; set; } = new();
}

public class ReportModel
{
    public List<CustomerGroup> CustomerGroups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Outer loop over customer groups.
        builder.Writeln("<<foreach [group in model.CustomerGroups]>>");
        builder.Writeln("Customer ID: <<[group.CustomerId]>>");
        builder.Writeln("Orders:");
        // Inner loop over orders belonging to the current customer.
        builder.Writeln("<<foreach [order in group.Orders]>>");
        builder.Writeln("- Order <<[order.OrderId]>>: <<[order.Product]>> - $<<[order.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data and group orders by CustomerId.
        // -----------------------------------------------------------------
        List<Order> orders = new()
        {
            new Order { OrderId = 1, CustomerId = 101, Product = "Laptop", Price = 1200.00 },
            new Order { OrderId = 2, CustomerId = 102, Product = "Smartphone", Price = 800.00 },
            new Order { OrderId = 3, CustomerId = 101, Product = "Mouse", Price = 25.50 },
            new Order { OrderId = 4, CustomerId = 103, Product = "Keyboard", Price = 45.00 },
            new Order { OrderId = 5, CustomerId = 102, Product = "Monitor", Price = 300.00 }
        };

        // Use LINQ GroupBy to create a collection of CustomerGroup objects.
        ReportModel model = new()
        {
            CustomerGroups = orders
                .GroupBy(o => o.CustomerId)
                .Select(g => new CustomerGroup
                {
                    CustomerId = g.Key,
                    Orders = g.ToList()
                })
                .ToList()
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);
    }
}
