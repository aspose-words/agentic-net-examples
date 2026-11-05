using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

public class ReportModel
{
    // Property used in a normal paragraph.
    public string Name { get; set; } = "";

    // Property used in a paragraph that contains only a tag.
    public string EmptyTag { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Paragraph with regular text and a tag.
        builder.Writeln("Hello <<[model.Name]>>!");

        // Paragraph that contains only a tag – should be removed if the tag resolves to empty.
        builder.Writeln("<<[model.EmptyTag]>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model with empty values.
        ReportModel model = new ReportModel
        {
            Name = "",          // Empty string – the tag will render nothing but the paragraph stays because it has other text.
            EmptyTag = ""       // Empty string – the whole paragraph should be removed.
        };

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        reportDoc.Save(resultPath);
    }
}
