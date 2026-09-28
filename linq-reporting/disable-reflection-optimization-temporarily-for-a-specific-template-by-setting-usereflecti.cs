using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // Create a simple template with a LINQ Reporting tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Sample data model.
        ReportModel model = new ReportModel { Name = "World" };

        // Temporarily disable reflection optimization.
        bool originalOptimization = ReportingEngine.UseReflectionOptimization;
        ReportingEngine.UseReflectionOptimization = false;
        try
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");
        }
        finally
        {
            // Restore original setting.
            ReportingEngine.UseReflectionOptimization = originalOptimization;
        }

        // Save the generated report.
        doc.Save(reportPath);
    }
}

// Public data model class.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
}
