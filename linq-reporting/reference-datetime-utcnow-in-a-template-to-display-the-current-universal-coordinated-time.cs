using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Returns the current UTC time when accessed.
    public DateTime CurrentUtc => DateTime.UtcNow;
}

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the example files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a simple paragraph that displays the current UTC time using a LINQ Reporting tag.
        builder.Writeln("Current UTC time: <<[model.CurrentUtc]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and generate the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare the root data object.
        ReportModel model = new ReportModel();

        // Create the reporting engine and build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}
