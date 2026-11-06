using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Valid property used in the template.
    public string Name { get; set; } = "John Doe";

    // No property named MissingProperty – this will trigger an inline error message.
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for full encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a temporary folder for the example files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add an empty paragraph that should be removed later.
        builder.Writeln(string.Empty);
        builder.Writeln(string.Empty);

        // Add a paragraph with a valid tag.
        builder.Writeln("Customer Name: <<[model.Name]>>");

        // Add a paragraph with an invalid tag to demonstrate inline error messages.
        builder.Writeln("Missing field: <<[model.MissingProperty]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportModel model = new ReportModel();

        // -----------------------------------------------------------------
        // Step 3: Configure the reporting engine with both options.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages;

        // Build the report. The returned flag indicates overall success.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
