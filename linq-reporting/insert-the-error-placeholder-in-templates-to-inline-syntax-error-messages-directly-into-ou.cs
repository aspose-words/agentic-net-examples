using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        Order order = new Order
        {
            CustomerName = "John Doe"
        };

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags, including a deliberate error and the <<error>> placeholder.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Invalid: <<[order.NonExisting]>>"); // This property does not exist.
        builder.Writeln("<<error>>"); // Placeholder for inline error messages.

        // Save the template to disk.
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back for report generation.
        Document doc = new Document(templatePath);

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The method returns false because there is a syntax error.
        bool success = engine.BuildReport(doc, order, "order");

        // Save the resulting document.
        string outputPath = "ReportWithInlineErrors.docx";
        doc.Save(outputPath);

        // Output the success flag.
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
    }
}

// Public data model class.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
}
