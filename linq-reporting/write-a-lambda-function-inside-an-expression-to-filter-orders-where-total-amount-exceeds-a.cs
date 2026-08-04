using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public decimal TotalAmount { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
    public decimal Threshold { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create the template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Header showing the threshold value.
        builder.Writeln("Orders with total greater than <<[model.Threshold]>>:");

        // Use a lambda expression inside the foreach tag to filter orders.
        builder.Writeln("<<foreach [order in model.Orders.Where(o => o.TotalAmount > model.Threshold)]>>");
        builder.Writeln("- <<[order.CustomerName]>>: $<<[order.TotalAmount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template for report generation.
        var doc = new Document(templatePath);

        // Step 3: Prepare sample data.
        var model = new ReportModel
        {
            Threshold = 100m,
            Orders = new List<Order>
            {
                new Order { CustomerName = "Alice", TotalAmount = 75m },
                new Order { CustomerName = "Bob",   TotalAmount = 150m },
                new Order { CustomerName = "Carol", TotalAmount = 200m },
                new Order { CustomerName = "Dave",  TotalAmount = 50m }
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Step 5: Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
