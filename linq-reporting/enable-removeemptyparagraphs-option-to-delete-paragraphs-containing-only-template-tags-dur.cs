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
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // Create a simple template document.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Paragraph with regular text.
        builder.Writeln("Customer Report");
        builder.Writeln();

        // Paragraph that contains only a LINQ Reporting tag.
        // This paragraph should be removed when RemoveEmptyParagraphs is enabled.
        builder.Writeln("<<[model.Name]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Sample data model.
        ReportModel model = new ReportModel
        {
            Name = "Acme Corporation"
        };

        // Configure the reporting engine with RemoveEmptyParagraphs option.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);
    }
}

// Simple public data model used by the template.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
}
