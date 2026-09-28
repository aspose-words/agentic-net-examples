using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write static content.
        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln("Total Amount: <<[order.Total]>>");

        // Conditional block: show discount label only when Total > 100.
        builder.Writeln("<<if [order.Total > 100]>>Discount Applied!<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        Order order = new()
        {
            Total = 150.00m // Change this value to test the condition.
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Public data model class.
public class Order
{
    // Initialize to avoid nullable warnings.
    public decimal Total { get; set; } = 0m;
}
