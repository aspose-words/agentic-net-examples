using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var order = new Order
        {
            CustomerName = "John Doe",
            TotalAmount = 12345.67m
        };

        // Create a template document with a LINQ Reporting tag.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Set the current thread culture to demonstrate locale‑specific formatting.
        // Change the culture identifier to test different locales (e.g., "de-DE", "ja-JP").
        Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Creates a simple Word template containing a LINQ Reporting expression.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a heading.
        builder.Writeln("Invoice");
        builder.Writeln();

        // Insert data fields.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total Amount: <<[order.TotalAmount]>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Public data model class used by the report.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal TotalAmount { get; set; }
}
