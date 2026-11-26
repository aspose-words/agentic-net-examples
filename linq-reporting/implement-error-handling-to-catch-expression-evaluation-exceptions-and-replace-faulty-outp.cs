using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public double Value { get; set; } = 10;
    public double Divisor { get; set; } = 0; // This will cause a divide‑by‑zero exception in the template
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create template document with a faulty expression
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Result: <<[model.Value / model.Divisor]>>"); // Division by zero will trigger an error
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare data model
        ReportModel model = new ReportModel();

        // Configure reporting engine to show inline error placeholders
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        string resultPath = Path.Combine(outputDir, "Result.docx");
        doc.Save(resultPath);

        // Output status (no interactive prompts)
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}
