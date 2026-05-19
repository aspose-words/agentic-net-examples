using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample data using plain .NET objects.
        ReportModel model = CreateReportModel();

        // Build a blank Word document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags into the template.
        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Customer: <<[c.CustomerName]>>");
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [o in c.Orders]>>");
        builder.Writeln("- Order ID: <<[o.OrderID]>>, Product: <<[o.Product]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The empty string means the root object itself is the data source.
        engine.BuildReport(doc, model, "");

        // Save the generated report.
        doc.Save("Report.docx");
    }

    // Creates a model that contains customers and their related orders.
    private static ReportModel CreateReportModel()
    {
        var customers = new List<Customer>
        {
            new Customer
            {
                CustomerID = 1,
                CustomerName = "John Doe",
                Orders = new List<Order>
                {
                    new Order { OrderID = 100, Product = "Laptop" },
                    new Order { OrderID = 101, Product = "Mouse" }
                }
            },
            new Customer
            {
                CustomerID = 2,
                CustomerName = "Jane Smith",
                Orders = new List<Order>
                {
                    new Order { OrderID = 102, Product = "Keyboard" }
                }
            }
        };

        return new ReportModel { Customers = customers };
    }
}

// Root wrapper class for the report.
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

// Customer class with a collection of orders.
public class Customer
{
    public int CustomerID { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public List<Order> Orders { get; set; } = new();
}

// Order class.
public class Order
{
    public int OrderID { get; set; }
    public string Product { get; set; } = string.Empty;
}
