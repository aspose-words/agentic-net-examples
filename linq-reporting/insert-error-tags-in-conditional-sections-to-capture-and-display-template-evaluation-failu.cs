using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample property; the template will reference a non‑existent property to trigger an error.
    public string Title { get; set; } = "Sample Report";
}

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // -----------------------------------------------------------------
        // Create the template document with a conditional section that
        // references a missing property. Inside the conditional we place
        // an <<error>> tag to capture evaluation failures.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("=== Report Start ===");
        builder.Writeln("<<if [model.MissingProperty]>>"); // This property does not exist.
        builder.Writeln("<<error>>");                     // Will display the evaluation error.
        builder.Writeln("<</if>>");
        builder.Writeln("=== Report End ===");

        // Save the template to disk (optional but satisfies the rule of saving before loading).
        template.Save(templatePath);

        // Load the template back for reporting.
        Document doc = new Document(templatePath);

        // Create the root data model.
        ReportModel model = new ReportModel();

        // Configure the reporting engine to emit inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The returned flag indicates overall success.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Template path: {templatePath}");
        Console.WriteLine($"Output path: {outputPath}");
    }
}
