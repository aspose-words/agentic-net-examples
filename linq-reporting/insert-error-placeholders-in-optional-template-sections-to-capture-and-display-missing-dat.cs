using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public string? OptionalField { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Report for <<[model.Name]>>");

        // Optional field with an <<error>> placeholder to capture missing data warnings.
        builder.Writeln("Optional Info: <<[model.OptionalField]>> <<error>>");

        // Footer.
        builder.Writeln("End of report.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // Create sample data. OptionalField is left null to trigger a warning.
        ReportModel model = new ReportModel
        {
            Name = "John Doe",
            OptionalField = null
        };

        // -------------------------------------------------
        // Build the report with inline error messages enabled.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {outputPath}");
    }
}
