using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        Order order = new()
        {
            Quantity = 7.9 // Non‑integer quantity to demonstrate casting.
        };

        // Create a template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Report for Order");
        // Use explicit cast (int) on Quantity when calling a method that expects an integer.
        builder.Writeln("<<[order.GetQuantityString((int)order.Quantity)]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Build the report.
        ReportingEngine.UseReflectionOptimization = true;
        ReportingEngine engine = new();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        string outputPath = Path.Combine("output", "report.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        doc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Public data model aligned with the template.
public class Order
{
    // Non‑integer quantity to require casting.
    public double Quantity { get; set; } = 0;

    // Method that expects an integer parameter.
    public string GetQuantityString(int qty)
    {
        return $"Quantity (as int): {qty}";
    }
}
