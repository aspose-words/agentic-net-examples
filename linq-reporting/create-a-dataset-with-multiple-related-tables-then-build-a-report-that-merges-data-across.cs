using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class OrderDetail
{
    public int OrderDetailID { get; set; }
    public string Product { get; set; } = "";
    public int Quantity { get; set; }
}

public class Order
{
    public int OrderID { get; set; }
    public DateTime OrderDate { get; set; }
    public List<OrderDetail> OrderDetails { get; set; } = new();
}

public class Customer
{
    public int CustomerID { get; set; }
    public string Name { get; set; } = "";
    public List<Order> Orders { get; set; } = new();
}

public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

class Program
{
    static void Main()
    {
        // Register code page provider for Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Bootstrap sample data.
        var model = new ReportModel();

        var alice = new Customer { CustomerID = 1, Name = "Alice" };
        var bob   = new Customer { CustomerID = 2, Name = "Bob" };

        var order100 = new Order { OrderID = 100, OrderDate = new DateTime(2023, 1, 15) };
        order100.OrderDetails.Add(new OrderDetail { OrderDetailID = 1000, Product = "Widget",     Quantity = 5 });
        order100.OrderDetails.Add(new OrderDetail { OrderDetailID = 1001, Product = "Gadget",     Quantity = 2 });

        var order101 = new Order { OrderID = 101, OrderDate = new DateTime(2023, 2, 5) };
        order101.OrderDetails.Add(new OrderDetail { OrderDetailID = 1002, Product = "Thingamajig", Quantity = 1 });

        var order102 = new Order { OrderID = 102, OrderDate = new DateTime(2023, 3, 12) };
        order102.OrderDetails.Add(new OrderDetail { OrderDetailID = 1003, Product = "Doohickey", Quantity = 3 });
        order102.OrderDetails.Add(new OrderDetail { OrderDetailID = 1004, Product = "Whatsit",   Quantity = 4 });

        alice.Orders.AddRange(new[] { order100, order101 });
        bob.Orders.Add(order102);

        model.Customers.AddRange(new[] { alice, bob });

        // Create the template document with LINQ Reporting tags.
        string templatePath = "ReportTemplate.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln();

        builder.Writeln("<<foreach [c in model.Customers]>>");
        builder.Writeln("Customer: <<[c.Name]>>");
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [o in c.Orders]>>");
        builder.Writeln("- Order ID: <<[o.OrderID]>>, Date: <<[o.OrderDate]>>");
        builder.Writeln("  Details:");
        builder.Writeln("  <<foreach [d in o.OrderDetails]>>");
        builder.Writeln("  * <<[d.Product]>> x <<[d.Quantity]>>");
        builder.Writeln("  <</foreach>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "ReportOutput.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine(success
            ? $"Report generated successfully: {Path.GetFullPath(outputPath)}"
            : "Report generation failed.");
    }
}
