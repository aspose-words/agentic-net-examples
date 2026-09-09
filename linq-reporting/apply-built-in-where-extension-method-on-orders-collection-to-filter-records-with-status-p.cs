using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public string Status { get; set; } = "";
    public string CustomerName { get; set; } = "";
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { Id = 1, Status = "Pending",   CustomerName = "Alice" },
                new Order { Id = 2, Status = "Shipped",   CustomerName = "Bob"   },
                new Order { Id = 3, Status = "Pending",   CustomerName = "Carol" },
                new Order { Id = 4, Status = "Delivered", CustomerName = "Dave"  }
            }
        };

        // Filter the collection to only pending orders using LINQ.
        model.Orders = model.Orders.Where(o => o.Status == "Pending").ToList();

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a simple LINQ Reporting tag that iterates over the Orders collection.
        builder.Writeln("Pending Orders:");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("  Id: <<[order.Id]>>  Customer: <<[order.CustomerName]>>  Status: <<[order.Status]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
