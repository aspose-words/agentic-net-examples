using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportApp
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words functionalities)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model
        Order order = new Order
        {
            Discount = 5.5m,
            Tax = 2.3m
        };

        // Create a template document programmatically
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template
        Document doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report
        string reportPath = "report.docx";
        doc.Save(reportPath);
    }

    // Generates a simple Word template containing LINQ Reporting tags
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Discount: <<[order.Discount]>>");
        builder.Writeln("Tax: <<[order.Tax]>>");
        builder.Writeln("Combined (Discount + Tax): <<[order.Combined]>>");

        doc.Save(filePath);
    }
}

// Data model used by the template
public class Order
{
    // Nullable decimal fields
    public decimal? Discount { get; set; }
    public decimal? Tax { get; set; }

    // Combined value using lifted addition operator (null propagates)
    public decimal? Combined => Discount + Tax;
}
