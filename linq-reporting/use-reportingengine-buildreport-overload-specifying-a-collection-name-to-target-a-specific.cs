using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data collection.
        List<Order> orders = new()
        {
            new Order { Id = 1, CustomerName = "Alice" },
            new Order { Id = 2, CustomerName = "Bob" },
            new Order { Id = 3, CustomerName = "Charlie" }
        };

        // Create a template document with LINQ Reporting tags.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>, Customer: <<[order.CustomerName]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, but ensures the document is fully created before building).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document doc = new(templatePath);

        // Build the report using the collection name overload.
        ReportingEngine engine = new();
        bool success = engine.BuildReport(doc, orders, "orders");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
