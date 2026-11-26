using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal Total { get; set; }
    // Additional properties can be added as needed.
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // Create the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Summary");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: <<[order.Total]>>");

        // Conditional block: display discount label only when Total exceeds 100.
        builder.Writeln("<<if [order.Total > 100]>>Discount Applied!<</if>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        Order order = new Order
        {
            CustomerName = "John Doe",
            Total = 150m // Exceeds the threshold, so the discount label will appear.
        };

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save(reportPath);
    }
}
