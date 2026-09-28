using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class InlineErrorMessageExample
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple template with a valid tag and an intentional reference error.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Correct tag.
        builder.Writeln("Hello, <<[model.Name]>>!");

        // Intentional reference error (property does not exist).
        builder.Writeln("This line has a reference error: <<[model.Unknown]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        ReportModel model = new()
        {
            Name = "John Doe"
        };

        // Configure the reporting engine to show inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        reportDoc.Save(resultPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }

    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }
}
