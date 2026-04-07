using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert sample data fields.
        builder.Writeln("Product: <<[order.ProductName]>>");
        builder.Writeln("Price: $<<[order.Price]>>");

        // Conditional block: show discount only when it is greater than zero.
        builder.Writeln("<<if [order.Discount > 0]>>Discount: <<[order.Discount]>>%<</if>>");

        // Build the report using a concrete data source.
        Order order = new Order
        {
            ProductName = "Wireless Mouse",
            Price = 29.99m,
            Discount = 15 // Set to 0 to hide the discount line.
        };

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, order, "order");

        // Save the generated report.
        template.Save("Report.docx");
    }
}

// Simple data model aligned with the template tags.
public class Order
{
    public string ProductName { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public int Discount { get; set; }
}
