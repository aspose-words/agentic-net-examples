using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create a template document with a LINQ Reporting tag.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        Model model = new Model { Name = "World" };

        // Configure the reporting engine with InlineErrorMessages option.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report and capture the success flag.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        reportDoc.Save(resultPath);

        // Output the success flag.
        Console.WriteLine($"Report build success: {success}");
    }
}

// Public data model class used by the template.
public class Model
{
    public string Name { get; set; } = string.Empty;
}
