using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a template document with LINQ Reporting tags
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Correct tag – will be replaced with the model's Name property
        builder.Writeln("Customer Name: <<[model.Name]>>");

        // Incorrect tag – malformed expression to trigger an inline error message
        builder.Writeln("Faulty Tag: <<[model.Missing>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // Load the template for report generation
        var doc = new Document(templatePath);

        // Prepare the data model
        var model = new ReportModel { Name = "John Doe" };

        // Configure the reporting engine to inline error messages
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report; the returned flag indicates overall success
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Output the result status
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to {outputPath}");
    }

    // Public data model used by the template
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }
}
