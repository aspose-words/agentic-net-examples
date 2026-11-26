using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Total amount of the order.
    public double Total { get; set; } = 0;
}

public class Program
{
    public static void Main()
    {
        // Prepare a simple template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write the total amount.
        builder.Writeln("Total: <<[order.Total]>>");

        // Compute discount using conditional operator.
        builder.Writeln("Discount: <<[order.Total > 500 ? order.Total * 0.1 : 0]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Create sample data.
        Order order = new Order { Total = 620.0 };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
