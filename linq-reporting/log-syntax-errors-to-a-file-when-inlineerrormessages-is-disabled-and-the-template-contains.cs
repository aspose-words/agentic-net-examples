using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Name { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");
        string logPath = Path.Combine(workDir, "error.log");

        // Create a template document with a valid and an invalid LINQ Reporting tag.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello <<[model.Name]>>!");               // Valid expression.
        builder.Writeln("This will cause an error: <<[model.Missing]>>"); // Invalid expression.
        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel();

        // Configure the reporting engine without InlineErrorMessages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // InlineErrorMessages is disabled.

        try
        {
            // Attempt to build the report. This will throw on syntax errors.
            bool success = engine.BuildReport(template, model, "model");

            // If no exception, save the generated document.
            if (success)
                template.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log the syntax error details to a file.
            File.WriteAllText(logPath, $"Report generation failed: {ex.Message}");
        }
    }
}
