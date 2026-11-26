using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template with an invalid expression (property does not exist).
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Header");
        builder.Writeln("<<[model.NonExistentProperty]>>"); // Invalid expression.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        Model model = new Model { Name = "Sample Name" };

        // Configure the reporting engine without InlineErrorMessages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // Path for the error log.
        string logPath = Path.Combine(outputDir, "ErrorLog.txt");

        try
        {
            // Attempt to build the report. This will throw because of the invalid expression.
            engine.BuildReport(reportDoc, model, "model");
        }
        catch (Exception ex)
        {
            // Log the error details to a file.
            File.WriteAllText(logPath, ex.ToString());
        }

        // Save the (possibly incomplete) report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Simple data model used for the report.
public class Model
{
    public string Name { get; set; } = string.Empty;
}
