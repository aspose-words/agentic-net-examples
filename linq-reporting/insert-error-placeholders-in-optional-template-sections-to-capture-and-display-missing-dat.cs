using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple data model.
        // -----------------------------------------------------------------
        var model = new ReportModel();

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Normal field – will be replaced with the actual value.
        builder.Writeln("Existing value: <<[model.Existing]>>");

        // Missing field – the property does not exist in the model.
        // The <<error>> placeholder will capture the warning when InlineErrorMessages is enabled.
        builder.Writeln("Missing value: <<[model.Missing]>> <<error>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            // Enable inline error messages so that <<error>> tags are populated.
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // BuildReport returns a bool indicating success when InlineErrorMessages is set.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save(reportPath);

        // Output simple status information.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}

// ---------------------------------------------------------------------
// Data model used by the report.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string Existing { get; set; } = "Present";
    // Note: No property named 'Missing' is defined on purpose.
}
