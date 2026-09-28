using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "ReportWithErrors.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a heading.
        builder.Writeln("LINQ Reporting Inline Error Example");
        builder.Writeln();

        // Correct tag – will be replaced with the model's Name.
        builder.Writeln("Customer Name: <<[model.Name]>>");

        // Incorrect tag – references a non‑existent property, will trigger an error.
        builder.Writeln("Missing Property: <<[model.MissingProperty]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Step 3: Prepare the data model.
        // -----------------------------------------------------------------
        var model = new SampleModel
        {
            Name = "John Doe"
        };

        // -----------------------------------------------------------------
        // Step 4: Configure the ReportingEngine to inline error messages.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The method returns false because there is an error.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Step 5: Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");
    }
}

// ---------------------------------------------------------------------
// Data model used by the template.
// ---------------------------------------------------------------------
public class SampleModel
{
    // Non‑nullable property with an initializer to avoid warnings.
    public string Name { get; set; } = string.Empty;
}
