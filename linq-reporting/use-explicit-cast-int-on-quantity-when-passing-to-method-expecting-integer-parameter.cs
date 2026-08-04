using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class OrderItem
{
    // Quantity is stored as double to demonstrate explicit casting in the template.
    public double Quantity { get; set; }
    public string Name { get; set; }

    // Method expects an integer parameter.
    public string GetQuantityMessage(int qty)
    {
        return $"Quantity is {qty}";
    }
}

public class ReportModel
{
    public List<OrderItem> Items { get; set; } = new List<OrderItem>();
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple title.
        builder.Writeln("Order Report");
        builder.Writeln();

        // Insert LINQ Reporting tags.
        // The foreach iterates over Items, and the method call casts Quantity to int.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Product: <<[item.Name]>>");
        builder.Writeln("Quantity Message: <<[item.GetQuantityMessage((int)item.Quantity)]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<OrderItem>
            {
                new OrderItem { Name = "Apple", Quantity = 5.7 },
                new OrderItem { Name = "Banana", Quantity = 3.0 },
                new OrderItem { Name = "Cherry", Quantity = 12.4 }
            }
        };

        // Build the report. No root name is needed because we reference members directly.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, null);

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
