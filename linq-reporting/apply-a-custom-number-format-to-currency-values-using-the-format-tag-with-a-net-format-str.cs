using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public decimal Amount { get; set; } = 0m;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document with a formatted currency tag
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Invoice Report");
        // Use ToString with format specifier instead of the unsupported -format switch
        builder.Writeln($"Total Amount: <<[order.Amount.ToString(\"C2\")]>>");
        doc.Save(templatePath);

        // Load the template
        var template = new Document(templatePath);

        // Prepare data model
        var order = new Order { Amount = 1234.56m };

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(template, order, "order");

        // Save the generated report
        const string outputPath = "report.docx";
        template.Save(outputPath);
    }
}
