using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        Order order = new Order
        {
            Total = 12345.67m
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a line with a custom number format for the currency value.
        // The format tag uses a .NET format string expression enclosed in double quotes.
        builder.Writeln("Order Total: <<[order.Total]:\"$#,##0.00\">>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template back (simulating a real‑world scenario where the template exists on disk).
        Document doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}

// Simple data model with a public property that will be referenced in the template.
public class Order
{
    public decimal Total { get; set; } = 0m;
}
