using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // Create a simple template with a LINQ Reporting tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello, <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Disable reflection optimization for this report generation.
        ReportingEngine.UseReflectionOptimization = false;

        // Prepare the data model.
        ReportModel model = new()
        {
            Name = "Aspose.Words User"
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine("Report generated at: " + reportPath);
    }
}

// Simple data model used by the template.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
}
