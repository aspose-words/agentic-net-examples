using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer
                {
                    Name = "Alice",
                    Orders = new List<Order>
                    {
                        new Order { Id = 1, ProductName = "Laptop", Quantity = 2 },
                        new Order { Id = 2, ProductName = "Mouse", Quantity = 5 }
                    }
                },
                new Customer
                {
                    Name = "Bob",
                    Orders = new List<Order>
                    {
                        new Order { Id = 3, ProductName = "Keyboard", Quantity = 1 },
                        new Order { Id = 4, ProductName = "Monitor", Quantity = 2 },
                        new Order { Id = 5, ProductName = "USB‑Cable", Quantity = 10 }
                    }
                }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [customer in Customers]>>");
        builder.Writeln("Customer: <<[customer.Name]>>");
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [order in customer.Orders]>>");
        builder.Writeln("- Id: <<[order.Id]>>, Product: <<[order.ProductName]>>, Qty: <<[order.Quantity]>>");
        builder.Writeln("<</foreach>>"); // End inner foreach (orders)
        builder.Writeln("<</foreach>>"); // End outer foreach (customers)

        // Save the template.
        builder.Document.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the model as the root data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model);

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Root model containing a collection of customers.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer entity with a name and a collection of orders.
public class Customer
{
    public string Name { get; set; } = "";
    public List<Order> Orders { get; set; } = new();
}

// Order entity with simple properties.
public class Order
{
    public int Id { get; set; }
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; }
}
