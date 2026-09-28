using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // The actual hyperlink target object (e.g., a Uri instance).
    public object HyperlinkTarget { get; set; }

    // Text that will be displayed for the hyperlink.
    public string LinkText { get; set; } = string.Empty;

    // Returns the string representation of the hyperlink target.
    public string HyperlinkTargetString => HyperlinkTarget?.ToString() ?? string.Empty;

    public ReportModel()
    {
        // Sample data: a Uri object as the hyperlink target.
        HyperlinkTarget = new Uri("https://example.com");
        LinkText = "Visit Example.com";
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a hyperlink using LINQ Reporting tags.
        // The URI expression uses the string-converted property.
        builder.Writeln("<<link [model.HyperlinkTargetString] [model.LinkText]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // 3. Prepare the data model.
        ReportModel model = new ReportModel();

        // 4. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // 5. Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        loadedTemplate.Save(resultPath);
    }
}
