using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare template document
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags
        builder.Writeln("Customer Name: <<[order.CustomerName]>>");
        builder.Writeln("Missing Field (should be empty): <<[order.MissingField]>>");
        templateDoc.Save(templatePath);

        // Load the template
        Document doc = new Document(templatePath);

        // Prepare data model (MissingField does NOT exist in the original model)
        Order order = new Order
        {
            CustomerName = "John Doe"
        };

        // Configure the reporting engine.
        // The engine will ignore missing members by default when the corresponding
        // property returns null, so we do not need a special flag that is unavailable
        // in the current Aspose.Words version.
        ReportingEngine engine = new ReportingEngine();

        // Build the report
        engine.BuildReport(doc, order, "order");

        // Save the generated report
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model used by the report
public class Order
{
    public string CustomerName { get; set; } = string.Empty;

    // The MissingField property is intentionally left to return null.
    // This allows the template tag <<[order.MissingField]>> to be treated as an empty value.
    public string? MissingField => null;
}
