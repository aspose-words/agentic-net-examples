using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportModel
{
    public string CustomerName { get; set; } = "Acme Corp";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with a correct tag and an intentional syntax error
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer: <<[model.CustomerName]>>");
        // Intentional syntax error: missing closing bracket for the expression
        builder.Writeln("Broken tag: <<[model.Missing>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel();

        // Configure the reporting engine to show inline error messages
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        // Output result information
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
