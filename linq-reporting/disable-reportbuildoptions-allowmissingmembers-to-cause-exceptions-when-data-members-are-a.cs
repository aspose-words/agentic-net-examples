using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Existing property – the template will reference a missing one to trigger an exception.
    public string ExistingProperty { get; set; } = "ExistingValue";
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
        // 1. Create a LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The template references a non‑existent member "MissingProperty".
        builder.Writeln("Report start");
        builder.Writeln("<<[model.MissingProperty]>>"); // This member does not exist in Model.
        builder.Writeln("Report end");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure ReportingEngine without AllowMissingMembers.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Disables AllowMissingMembers.

        // -----------------------------------------------------------------
        // 4. Build the report and handle the expected exception.
        // -----------------------------------------------------------------
        Model model = new Model();

        try
        {
            // This call should throw because "MissingProperty" is not present.
            engine.BuildReport(reportDoc, model, "model");
            // If no exception, save the generated report.
            reportDoc.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            // Expected path: report generation fails due to missing member.
            Console.WriteLine($"Exception during report generation: {ex.Message}");
        }
    }
}
