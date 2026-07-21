using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Name { get; set; } = "Sample Name";
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template, generated report and error log.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");
        string errorLogPath = Path.Combine(outputDir, "error.log");

        // -----------------------------------------------------------------
        // 1. Create a template document with an invalid LINQ Reporting tag.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Valid text.
        builder.Writeln("Report Header");
        // Invalid tag – missing one closing '>' character.
        builder.Writeln("<<[model.Name]>");
        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (as required by the workflow).
        // -----------------------------------------------------------------
        var loadedDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        var model = new ReportModel();

        // -----------------------------------------------------------------
        // 4. Build the report with InlineErrorMessages disabled.
        //    Capture any syntax errors and write them to a log file.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // InlineErrorMessages disabled.

        try
        {
            // This will throw an exception because the template contains a syntax error.
            engine.BuildReport(loadedDoc, model, "model");
        }
        catch (Exception ex)
        {
            // Log the exception message (and optionally the stack trace) to a file.
            File.WriteAllText(errorLogPath, $"Error building report:{Environment.NewLine}{ex.Message}");
        }

        // Save the (unmodified) document so we have an output file.
        loadedDoc.Save(reportPath);
    }
}
