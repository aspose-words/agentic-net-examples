using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple data model.
        var model = new OrderModel
        {
            Quantity = 7.5 // Non‑integer quantity to demonstrate explicit casting.
        };

        // Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Show the original quantity.
        builder.Writeln("Quantity: <<[model.Quantity]>>");

        // Call a method that expects an int, casting the double quantity explicitly.
        builder.Writeln("Discount (as %): <<[model.GetDiscount((int)model.Quantity)]>>");

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save("Report.docx");
    }
}

// Public data model required by the template.
public class OrderModel
{
    // Non‑nullable property initialized to avoid warnings.
    public double Quantity { get; set; } = 0;

    // Method that expects an integer parameter.
    public int GetDiscount(int quantity)
    {
        // Simple discount calculation: 10% per item.
        return quantity * 10;
    }
}
