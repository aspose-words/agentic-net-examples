using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public string Product { get; set; } = "";
    public int Quantity { get; set; }
}

public class Customer
{
    public string Name { get; set; } = "";
    public List<Order> Orders { get; set; } = new();
}

public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
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
                        new Order { Id = 1, Product = "Apple", Quantity = 5 },
                        new Order { Id = 2, Product = "Banana", Quantity = 3 }
                    }
                },
                new Customer
                {
                    Name = "Bob",
                    Orders = new List<Order>
                    {
                        new Order { Id = 3, Product = "Carrot", Quantity = 7 },
                        new Order { Id = 4, Product = "Dates", Quantity = 2 }
                    }
                }
            }
        };

        // Create a blank template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Outer foreach – iterate over customers.
        builder.Writeln("<<foreach [c in model.Customers]>>");
        builder.Writeln("Customer: <<[c.Name]>>");
        builder.Writeln();

        // Inner foreach – iterate over orders of the current customer.
        builder.Writeln("<<foreach [o in c.Orders]>>");
        builder.Writeln("- <<[o.Product]>> (Qty: <<[o.Quantity]>>)");
        builder.Writeln("<</foreach>>");

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("GroupedReport.docx");
    }
}
