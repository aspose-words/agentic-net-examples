using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class DecimalExtensions
{
    public static string ToCurrencyString(this decimal amount) => $"${amount:F2}";
}

public class Order
{
    public string CustomerName { get; set; } = "John Doe";
    public decimal Amount { get; set; } = 1234.56m;

    // Helper property that uses the extension method for formatting.
    public string FormattedAmount => Amount.ToCurrencyString();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words for some encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount: <<[order.FormattedAmount]>>");
        builder.Document.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var order = new Order();

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
