using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public DateTime OrderDate { get; set; }
    public decimal Amount { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
    // Collection pre‑filtered to the last month – computed in code, not in the template.
    public List<Order> RecentOrders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample data.
        // -----------------------------------------------------------------
        var orders = new List<Order>
        {
            new Order { CustomerName = "Alice", OrderDate = DateTime.Now.AddDays(-5), Amount = 120.50m },
            new Order { CustomerName = "Bob",   OrderDate = DateTime.Now.AddDays(-20), Amount = 75.00m },
            new Order { CustomerName = "Carol", OrderDate = DateTime.Now.AddMonths(-2), Amount = 200.00m } // older than a month
        };

        // Filter orders that fall within the last month.
        var recentOrders = orders
            .Where(o => o.OrderDate >= DateTime.Now.AddMonths(-1))
            .ToList();

        var model = new ReportModel
        {
            Orders = orders,
            RecentOrders = recentOrders
        };

        // -----------------------------------------------------------------
        // 2. Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        var docTemplate = new Document();
        var builder = new DocumentBuilder(docTemplate);

        builder.Writeln("Orders placed within the last month:");
        // Use the pre‑filtered collection in the foreach tag.
        builder.Writeln("<<foreach [order in model.RecentOrders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        docTemplate.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
