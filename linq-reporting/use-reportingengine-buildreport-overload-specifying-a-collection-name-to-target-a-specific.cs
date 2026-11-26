using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int Id { get; set; }
    public string CustomerName { get; set; } = "";
}

public class ReportData
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var data = new ReportData
        {
            Orders = new List<Order>
            {
                new Order { Id = 1, CustomerName = "Alice" },
                new Order { Id = 2, CustomerName = "Bob" }
            }
        };

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the collection name overload.
        // Pass the collection itself (data.Orders) and the collection name "Orders".
        var engine = new ReportingEngine();
        engine.BuildReport(doc, data.Orders, "Orders");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
