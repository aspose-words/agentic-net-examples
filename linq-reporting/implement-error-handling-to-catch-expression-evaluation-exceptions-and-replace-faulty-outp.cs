using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize properties to avoid nullable warnings.
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the example files.
        string outputDir = "Output";
        System.IO.Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");
        // This tag references a non‑existent member and will cause an evaluation error.
        builder.Writeln("Missing: <<[model.Unknown]>>");

        // Save the template to disk.
        string templatePath = System.IO.Path.Combine(outputDir, "Template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template document.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // Step 3: Prepare the data source.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel();

        // -----------------------------------------------------------------
        // Step 4: Configure the ReportingEngine with inline error handling.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            // InlineErrorMessages makes the engine insert error messages instead of throwing.
            Options = ReportBuildOptions.InlineErrorMessages,
            // This text will replace any expression that cannot be evaluated.
            MissingMemberMessage = "N/A"
        };

        // Build the report. The boolean indicates whether parsing succeeded.
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // Step 5: Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = System.IO.Path.Combine(outputDir, "Report.docx");
        loadedTemplate.Save(reportPath);

        // Output the result status.
        Console.WriteLine($"Report generation successful: {success}");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
