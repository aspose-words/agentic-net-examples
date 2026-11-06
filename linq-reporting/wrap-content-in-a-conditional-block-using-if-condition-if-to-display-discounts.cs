using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "John Doe",
            Discount = 15.5m // Set to 0 to test the false branch.
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("<<if [order.Discount > 0]>>Discount: <<[order.Discount]>>%<</if>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Public data model aligned with the template.
public class Order
{
    public string CustomerName { get; set; } = "";
    public decimal Discount { get; set; }
}
