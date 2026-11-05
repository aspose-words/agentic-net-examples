using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public double Quantity { get; set; }

    public string GetQuantityMessage(int qty)
    {
        return $"Quantity is {qty}";
    }
}

public class Program
{
    public static void Main()
    {
        // Create sample data.
        var order = new Order
        {
            CustomerName = "Acme Corp",
            Quantity = 7.9 // Non‑integer quantity.
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Message: <<[order.GetQuantityMessage((int)order.Quantity)]>>");
        builder.Document.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
