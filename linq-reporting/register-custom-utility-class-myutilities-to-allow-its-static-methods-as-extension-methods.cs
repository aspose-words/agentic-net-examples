using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    public static string FormatDate(DateTime date) => date.ToString("yyyy-MM-dd");
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public DateTime OrderDate { get; set; }

    // Helper property that uses the utility method – no registration required.
    public string FormattedOrderDate => MyUtilities.FormatDate(OrderDate);
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var order = new Order
        {
            CustomerName = "Acme Corp",
            OrderDate = new DateTime(2023, 12, 15)
        };

        // Create a template document programmatically
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order Date: <<[order.FormattedOrderDate]>>");
        doc.Save(templatePath);

        // Load the template (optional, can reuse the same document)
        var template = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(template, order, "order");

        // Save the generated report
        const string outputPath = "Report.docx";
        template.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
