using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public DateTime OrderDate { get; set; }
    public string CustomerName { get; set; } = string.Empty;
}

public class Model
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings Aspose might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Prepare sample data ----------
        var allOrders = new List<Order>
        {
            new Order { Id = 1, OrderDate = DateTime.Today.AddDays(-5),  CustomerName = "Alice" },
            new Order { Id = 2, OrderDate = DateTime.Today.AddDays(-20), CustomerName = "Bob"   },
            new Order { Id = 3, OrderDate = DateTime.Today.AddMonths(-2), CustomerName = "Carol" },
            new Order { Id = 4, OrderDate = DateTime.Today.AddDays(-30), CustomerName = "Dave"  },
            new Order { Id = 5, OrderDate = DateTime.Today.AddDays(-1),  CustomerName = "Eve"   }
        };

        // Filter orders placed within the last month using a lambda expression.
        DateTime now = DateTime.Today;
        DateTime monthAgo = now.AddMonths(-1);
        var recentOrders = allOrders
            .Where(o => o.OrderDate >= monthAgo && o.OrderDate < now)
            .ToList();

        var model = new Model { Orders = recentOrders };

        // ---------- Create the LINQ Reporting template ----------
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders placed within the last month:");
        // Use the root name "model" in the tags.
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Id: <<[order.Id]>>, Date: <<[order.OrderDate]>>, Customer: <<[order.CustomerName]>>");
        builder.Writeln("<</foreach>>");

        // Save the template before building the report.
        templateDoc.Save(templatePath);

        // ---------- Load the template and build the report ----------
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // ---------- Save the final report ----------
        string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}
