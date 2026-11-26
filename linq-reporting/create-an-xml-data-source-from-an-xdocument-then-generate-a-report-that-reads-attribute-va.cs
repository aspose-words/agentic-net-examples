using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string Id { get; set; } = "";
    public string Customer { get; set; } = "";
    public decimal Total { get; set; }
}

public class OrdersModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new OrdersModel
        {
            Orders =
            {
                new Order { Id = "1", Customer = "Alice", Total = 123.45m },
                new Order { Id = "2", Customer = "Bob",   Total =  67.89m }
            }
        };

        // Create a Word template programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Orders Report");
        builder.Writeln("==============");
        builder.Writeln();
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order Id: <<[order.Id]>>, Customer: <<[order.Customer]>>, Total: <<[order.Total]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
