using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create the LINQ Reporting template.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Outer loop: customers.
        builder.Writeln("<<foreach [customer in model.Customers]>>");
        builder.Writeln("Customer: <<[customer.Name]>>");
        builder.Writeln("Orders:");
        // Inner loop: orders of the current customer.
        builder.Writeln("<<foreach [order in customer.Orders]>>");
        builder.Writeln("- <<[order.ProductName]>>: <<[order.Quantity]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template for report generation.
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            Customers = new()
            {
                new Customer
                {
                    Name = "Alice",
                    Orders = new()
                    {
                        new Order { ProductName = "Apple", Quantity = 5 },
                        new Order { ProductName = "Banana", Quantity = 3 }
                    }
                },
                new Customer
                {
                    Name = "Bob",
                    Orders = new()
                    {
                        new Order { ProductName = "Carrot", Quantity = 7 },
                        new Order { ProductName = "Dates", Quantity = 2 }
                    }
                }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Wrapper class that serves as the root data source for the template.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer entity with a collection of orders.
public class Customer
{
    public string Name { get; set; } = string.Empty;
    public List<Order> Orders { get; set; } = new();
}

// Simple order entity.
public class Order
{
    public string ProductName { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
