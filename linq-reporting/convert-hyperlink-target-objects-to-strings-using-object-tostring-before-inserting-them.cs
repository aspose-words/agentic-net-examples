using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Hyperlink target stored as an object (e.g., Uri)
    public object HyperlinkTarget { get; set; } = new Uri("https://example.com");

    // Display text for the link
    public string LinkText { get; set; } = "Visit Example";

    // Convert the object to string before inserting into the template
    public string HyperlinkTargetString => HyperlinkTarget?.ToString() ?? string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create template document
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a link tag that uses the string-converted property
        builder.Writeln("Link: <<link [model.HyperlinkTargetString] [model.LinkText]>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare model data
        ReportModel model = new ReportModel();

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(outputDir, "report.docx");
        doc.Save(outputPath);
    }
}
