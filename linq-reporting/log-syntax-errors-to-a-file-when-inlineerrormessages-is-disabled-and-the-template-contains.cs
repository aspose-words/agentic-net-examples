using System;
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
        var model = new SampleModel
        {
            Name = "John Doe"
        };

        // Create a template document with a valid and an invalid expression.
        var templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Configure the reporting engine without InlineErrorMessages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // InlineErrorMessages disabled.

        bool success;
        try
        {
            // Build the report.
            success = engine.BuildReport(doc, model, "model");
        }
        catch (Exception ex)
        {
            // When InlineErrorMessages is disabled, errors are thrown as exceptions.
            success = false;
            LogError(ex);
        }

        // If the build failed, log syntax errors to a file.
        if (!success)
        {
            var logPath = "error.log";
            var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Report generation failed due to syntax errors in the template.";
            File.WriteAllText(logPath, logMessage);
        }
        else
        {
            // Save the generated report when successful.
            doc.Save("output.docx");
        }
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Valid expression.
        builder.Writeln("Customer Name: <<[model.Name]>>");

        // Invalid expression (syntax error – property does not exist).
        builder.Writeln("Invalid Field: <<[model.NonExistent]>>");

        doc.Save(path);
    }

    private static void LogError(Exception ex)
    {
        var errorLogPath = "exception.log";
        var sb = new StringBuilder();
        sb.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Exception during report generation:");
        sb.AppendLine(ex.ToString());
        File.WriteAllText(errorLogPath, sb.ToString());
    }
}

public class SampleModel
{
    public string Name { get; set; } = string.Empty;
}
