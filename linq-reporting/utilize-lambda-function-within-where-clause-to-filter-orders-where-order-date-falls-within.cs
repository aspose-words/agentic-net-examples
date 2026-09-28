using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public DateTime OrderDate { get; set; }
    public string CustomerName { get; set; } = string.Empty;
    public int OrderId { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample orders with various dates.
        var allOrders = new List<Order>
        {
            new Order { OrderId = 1, CustomerName = "Alice",   OrderDate = DateTime.Now.AddMonths(-2) },
            new Order { OrderId = 2, CustomerName = "Bob",     OrderDate = DateTime.Now.AddMonths(-1).AddDays(-10) },
            new Order { OrderId = 3, CustomerName = "Charlie", OrderDate = DateTime.Now.AddMonths(-1).AddDays(-2) },
            new Order { OrderId = 4, CustomerName = "Diana",   OrderDate = DateTime.Now.AddDays(-5) },
            new Order { OrderId = 5, CustomerName = "Eve",     OrderDate = DateTime.Now }
        };

        // Determine the first day of the previous month and the first day of the current month.
        DateTime now = DateTime.Now;
        DateTime firstDayCurrentMonth = new DateTime(now.Year, now.Month, 1);
        DateTime firstDayLastMonth = firstDayCurrentMonth.AddMonths(-1);

        // Filter orders whose OrderDate falls within last month.
        var filteredOrders = allOrders
            .Where(o => o.OrderDate >= firstDayLastMonth && o.OrderDate < firstDayCurrentMonth)
            .ToList();

        // Build the data model.
        var model = new ReportModel { Orders = filteredOrders };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Orders from Last Month:");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
