using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Initialize non‑nullable reference types to avoid CS8618 warnings.
    public string CustomerName { get; set; } = "";
    public double Total { get; set; }
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert simple fields.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: <<[order.Total]>>");

        // Conditional block – the label appears only when Total > 100.
        builder.Writeln("<<if [order.Total > 100]>>Discount Applied<</if>>");

        // 2. Prepare sample data.
        Order order = new Order
        {
            CustomerName = "John Doe",
            Total = 150.0 // Change this value to test the condition.
        };

        // 3. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the tag prefix used in the template ("order").
        engine.BuildReport(template, order, "order");

        // 4. Save the generated document.
        const string outputPath = "Report.docx";
        template.Save(outputPath);

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Report generated and saved to '{outputPath}'.");
    }
}
