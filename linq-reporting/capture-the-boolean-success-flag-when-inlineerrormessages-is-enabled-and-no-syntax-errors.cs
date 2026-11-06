using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize property to avoid nullable warnings.
    public string Name { get; set; } = "World";
}

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "LinqReportingExample");
        Directory.CreateDirectory(workDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // -----------------------------------------------------------------
        // 1. Build the template document with a simple LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!"); // Tag references the root object named "model".
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (required before building the report).
        // -----------------------------------------------------------------
        Document doc = new(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new();

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine with InlineErrorMessages option.
        // -----------------------------------------------------------------
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // -----------------------------------------------------------------
        // 5. Build the report and capture the success flag.
        // -----------------------------------------------------------------
        bool success = engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // Output the result.
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Generated document saved to: {outputPath}");
    }
}
