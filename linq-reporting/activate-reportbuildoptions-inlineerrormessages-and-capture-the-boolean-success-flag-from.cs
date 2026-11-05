using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample property used in the template.
    public string Name { get; set; } = "World";
}

public class Program
{
    public static void Main()
    {
        // Prepare paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // Create a simple template with a LINQ Reporting tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Configure the reporting engine with InlineErrorMessages option.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report and capture the success flag.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(resultPath);

        // Output the result.
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}
