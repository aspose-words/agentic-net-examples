using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "John Doe",
            OrderDate = new DateTime(2023, 7, 15)
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order Date: <<[Utility.FormatDate(order.OrderDate)]>>");

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Utility));

        // Build the report using the root object "order".
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        var outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }
}

// Sample data model.
public class Order
{
    public string CustomerName { get; set; } = "";
    public DateTime OrderDate { get; set; }
}

// Utility class with a static method accessed from the template.
public static class Utility
{
    public static string FormatDate(DateTime date) => date.ToString("yyyy-MM-dd");
}
