using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // Create a simple DOCX template with a LINQ Reporting tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        templateDoc.Save(templatePath);

        // Load the template from disk.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        Order order = new Order
        {
            CustomerName = "John Doe"
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save(outputPath);
    }
}

// Public data model class.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
}
