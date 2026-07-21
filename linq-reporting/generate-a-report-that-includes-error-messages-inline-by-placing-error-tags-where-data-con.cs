using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample numeric property.
    public int Age { get; set; } = 30;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string outputPath = "ReportOutput.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Normal tag – will be replaced with the value of Age.
        builder.Writeln("Customer age: <<[model.Age]>>");

        // Tag that will cause a runtime error (division by zero).
        // With InlineErrorMessages enabled, the engine will insert an <<error>> tag here.
        builder.Writeln("Age divided by zero (will cause error): <<[model.Age] / 0>>");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportModel model = new ReportModel();

        ReportingEngine engine = new ReportingEngine
        {
            // Enable inline error messages.
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // BuildReport returns true only when InlineErrorMessages is set.
        bool success = engine.BuildReport(reportDoc, model, "model");

        Console.WriteLine($"BuildReport succeeded: {success}");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
        Console.WriteLine($"Report saved to: {Path.GetFullPath(outputPath)}");
    }
}
