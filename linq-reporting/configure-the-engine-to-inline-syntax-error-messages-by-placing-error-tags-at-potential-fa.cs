using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "John Doe";
    // Additional properties can be added here.
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for potential encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare paths.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // Create a template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write tags: a valid tag and an intentionally invalid tag to trigger an inline error.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Missing property: <<[order.NonExisting]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        Order order = new Order();

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The returned flag indicates overall success.
        bool success = engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save(outputPath);

        // Optionally, write the result status to the console.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {outputPath}");
    }
}
