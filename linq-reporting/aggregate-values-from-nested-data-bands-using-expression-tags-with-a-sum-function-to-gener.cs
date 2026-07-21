using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Service
{
    public string Name { get; set; } = "";
    public decimal Cost { get; set; }
}

public class Order
{
    public int OrderId { get; set; }
    public string CustomerName { get; set; } = "";
    public List<Service> Services { get; set; } = new();

    // Calculated total for the order
    public decimal Total => Services?.Sum(s => s.Cost) ?? 0m;
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();

    // Calculated grand total for all orders
    public decimal GrandTotal => Orders?.Sum(o => o.Total) ?? 0m;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1,
                    CustomerName = "Alice",
                    Services = new List<Service>
                    {
                        new Service { Name = "Consulting", Cost = 150.00m },
                        new Service { Name = "Support", Cost = 75.00m }
                    }
                },
                new Order
                {
                    OrderId = 2,
                    CustomerName = "Bob",
                    Services = new List<Service>
                    {
                        new Service { Name = "Implementation", Cost = 300.00m },
                        new Service { Name = "Training", Cost = 120.00m },
                        new Service { Name = "Maintenance", Cost = 80.00m }
                    }
                }
            }
        };

        // Create template document
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Iterate orders
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Services:");
        // Iterate services
        builder.Writeln("<<foreach [svc in order.Services]>>");
        builder.Writeln("- <<[svc.Name]>> : $<<[svc.Cost]>>");
        builder.Writeln("<</foreach>>");
        // Order total using calculated property
        builder.Writeln("Order Total: $<<[order.Total]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Overall summary table
        builder.Writeln("Overall Summary");
        var table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Order ID");
        builder.InsertCell();
        builder.Writeln("Customer");
        builder.InsertCell();
        builder.Writeln("Total Cost");
        builder.EndRow();

        // Populate summary rows using foreach over Orders
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.OrderId]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.CustomerName]>>");
        builder.InsertCell();
        builder.Writeln("$<<[order.Total]>>");
        builder.EndRow();
        builder.Writeln("<</foreach>>");

        // Grand total row
        builder.InsertCell();
        builder.Writeln("Grand Total");
        builder.InsertCell();
        builder.Writeln("");
        builder.InsertCell();
        builder.Writeln("$<<[model.GrandTotal]>>");
        builder.EndRow();

        builder.EndTable();

        // Save the template
        doc.Save(templatePath);

        // Load the template for reporting
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        var outputPath = "OrdersReport.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
