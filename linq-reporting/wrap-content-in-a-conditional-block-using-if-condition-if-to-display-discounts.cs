using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public int OrderId { get; set; } = 0;
    public double Discount { get; set; } = 0.0;
}

public class Program
{
    public static void Main()
    {
        // Create a sample order with a discount.
        Order order = new Order
        {
            OrderId = 12345,
            Discount = 15.0 // Set to 0 to test the condition where no discount is shown.
        };

        // Create a blank Word document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert order information.
        builder.Writeln("Order ID: <<[order.OrderId]>>");

        // Conditional block: display the discount only when it is greater than zero.
        builder.Writeln("<<if [order.Discount > 0]>>Discount: <<[order.Discount]>>%<</if>>");

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save("ReportWithDiscount.docx");
    }
}
