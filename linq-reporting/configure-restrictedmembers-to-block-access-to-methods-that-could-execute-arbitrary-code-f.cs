using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var model = new SampleModel();

        // Create a template document programmatically.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Attempt to execute unsafe code: <<[model.Evil()]>>");
        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();

        // Enable inline error messages to capture any security violations in the output.
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        template.Save(outputPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Generated document saved to: {Path.GetFullPath(outputPath)}");
    }
}

// Sample data model with a method that attempts to start an external process.
// The actual process start is omitted to comply with security restrictions.
public class SampleModel
{
    // This method would normally launch the calculator, but the call is omitted for safety.
    public string Evil()
    {
        // Simulate a blocked operation.
        return "Process start blocked for security.";
    }
}
