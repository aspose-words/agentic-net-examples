using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public int Value { get; set; } = 10;
    public int Divisor { get; set; } = 0; // Will cause divide by zero
}

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Expression that will throw an exception (division by zero).
        builder.Writeln("Result of division: <<[model.Value / model.Divisor]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 2. Load the template for report generation.
        Document reportDoc = new(templatePath);

        // 3. Prepare the data model.
        ReportModel model = new();

        // 4. Configure the reporting engine.
        ReportingEngine engine = new();
        // Enable inline error messages; the engine will replace faulty output with placeholder text.
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // 5. Build the report. The method returns false if there were evaluation errors.
        bool success;
        try
        {
            success = engine.BuildReport(reportDoc, model, "model");
        }
        catch (Exception ex)
        {
            // Unexpected exception handling – treat as failure.
            Console.WriteLine($"Unexpected error: {ex.Message}");
            success = false;
        }

        // 6. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        // 7. Output the result status.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
