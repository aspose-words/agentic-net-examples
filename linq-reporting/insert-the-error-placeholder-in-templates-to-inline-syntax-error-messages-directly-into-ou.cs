using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample property used in the template.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Correct tag – will be replaced with the actual value.
        builder.Writeln("Customer: <<[model.Name]>>");

        // Incorrect tag – references a non‑existent member.
        // The <<error>> placeholder will be replaced with the inline error message.
        builder.Writeln("Invalid reference: <<[model.Unknown]>> <<error>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        Model data = new Model { Name = "John Doe" };

        // -------------------------------------------------
        // 3. Build the report with inline error messages enabled.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // BuildReport returns a bool indicating success when InlineErrorMessages is set.
        bool success = engine.BuildReport(reportDoc, data, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);

        // Output the result status to the console (no user interaction required).
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
        Console.WriteLine($"Template: {templatePath}");
        Console.WriteLine($"Report:   {reportPath}");
    }
}
