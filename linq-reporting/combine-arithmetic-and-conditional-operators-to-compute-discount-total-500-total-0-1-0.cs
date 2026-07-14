using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public double Total { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert template tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template: iterate over Orders and display total and discount.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order Total: <<[order.Total]>>");
        builder.Writeln("Discount: <<[order.Total > 500 ? order.Total * 0.1 : 0]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel();
        model.Orders.Add(new Order { Total = 750 });
        model.Orders.Add(new Order { Total = 420 });
        model.Orders.Add(new Order { Total = 1200 });

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("DiscountReport.docx");
    }
}
