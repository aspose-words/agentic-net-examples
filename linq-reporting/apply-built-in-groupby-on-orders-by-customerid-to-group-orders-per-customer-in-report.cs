using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int OrderId { get; set; }
    public int CustomerId { get; set; }
    public double Amount { get; set; }
}

public class CustomerGroup
{
    public int CustomerId { get; set; }
    public List<Order> Orders { get; set; } = new();
}

public class ReportModel
{
    public List<CustomerGroup> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create the template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Orders Report");
        // Iterate over each customer group.
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Customer ID: <<[group.CustomerId]>>");
        // Iterate over orders within the current group.
        builder.Writeln("<<foreach [order in group.Orders]>>");
        builder.Writeln("- Order ID: <<[order.OrderId]>>  Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk as required by the lifecycle rule.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template for report generation.
        var doc = new Document(templatePath);

        // 3. Prepare sample data and group orders by CustomerId.
        var orders = new List<Order>
        {
            new Order { OrderId = 1, CustomerId = 100, Amount = 250.0 },
            new Order { OrderId = 2, CustomerId = 101, Amount = 150.5 },
            new Order { OrderId = 3, CustomerId = 100, Amount = 99.99 },
            new Order { OrderId = 4, CustomerId = 102, Amount = 300.0 },
            new Order { OrderId = 5, CustomerId = 101, Amount = 45.75 }
        };

        var model = new ReportModel
        {
            Groups = orders
                .GroupBy(o => o.CustomerId)
                .Select(g => new CustomerGroup
                {
                    CustomerId = g.Key,
                    Orders = g.ToList()
                })
                .ToList()
        };

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model);

        // 5. Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
