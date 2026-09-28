using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template, result and log files
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "result.docx");
        string logPath = Path.Combine(outputDir, "errors.log");

        // -------------------------------------------------
        // Create a template document with an invalid expression
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report Header");
        // Invalid expression: property 'MissingProperty' does not exist in the model
        builder.Writeln("<<[model.MissingProperty]>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model (does NOT contain MissingProperty)
        ReportModel model = new ReportModel
        {
            Title = "Sample Report"
        };

        // -------------------------------------------------
        // Build the report without InlineErrorMessages option
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        try
        {
            // InlineErrorMessages flag is NOT set, so syntax errors will raise an exception
            engine.BuildReport(reportDoc, model, "model");

            // If no exception, save the generated document
            reportDoc.Save(resultPath);
        }
        catch (Exception ex)
        {
            // Log the syntax error details to a file
            File.WriteAllText(logPath, ex.ToString());

            // Optionally, still save the partially generated document
            reportDoc.Save(resultPath);
        }
    }

    // Simple public data model used by the template
    public class ReportModel
    {
        // Property that exists (used for demonstration)
        public string Title { get; set; } = string.Empty;
    }
}
