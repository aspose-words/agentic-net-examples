using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        doc.Save(templatePath);

        // Load the template.
        var template = new Document(templatePath);

        // Prepare the data model.
        var order = new Order { CustomerName = "Acme Corporation" };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, order, "order");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        template.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
