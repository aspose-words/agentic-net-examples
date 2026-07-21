using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public DateTime OrderDate { get; set; }
    public string CustomerName { get; set; } = string.Empty;
}

public class ReportModel
{
    // Full list of orders (sample data)
    public List<Order> Orders { get; set; } = new();

    // Orders filtered to the last month using a lambda in a Where clause
    public List<Order> FilteredOrders => Orders
        .Where(o => o.OrderDate >= DateTime.Now.AddMonths(-1))
        .ToList();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders placed in the last month:");
        // foreach over the filtered collection defined in the model
        builder.Writeln("<<foreach [order in model.FilteredOrders]>>");
        builder.Writeln("- Order ID: <<[order.Id]>>, Date: <<[order.OrderDate]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation
        // -------------------------------------------------
        var doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare sample data
        // -------------------------------------------------
        var model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { Id = 1, OrderDate = DateTime.Now.AddDays(-5),  CustomerName = "Alice" },
                new Order { Id = 2, OrderDate = DateTime.Now.AddDays(-20), CustomerName = "Bob"   },
                new Order { Id = 3, OrderDate = DateTime.Now.AddMonths(-2), CustomerName = "Carol" },
                new Order { Id = 4, OrderDate = DateTime.Now.AddDays(-2),  CustomerName = "Dave"  }
            }
        };

        // -------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report
        // -------------------------------------------------
        doc.Save(reportPath);
    }
}
