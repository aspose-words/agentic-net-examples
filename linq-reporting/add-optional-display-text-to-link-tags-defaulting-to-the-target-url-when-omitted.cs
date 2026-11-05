using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinkModel
{
    // URL to link to – required.
    public string Url { get; set; } = "https://example.com";

    // Optional display text – can be null or empty.
    public string? LinkText { get; set; } = "Visit Example";
}

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and result.
        string templatePath = Path.Combine(outputDir, "LinkTemplate.docx");
        string resultPath = Path.Combine(outputDir, "LinkReport.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Paragraph with link that provides explicit display text.
        builder.Writeln("Link with custom text:");
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Paragraph with link that omits display text – engine should use URL as text.
        builder.Writeln("Link with default text (URL will be shown):");
        builder.Writeln("<<link [model.Url]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data: first instance has custom text, second will use default.
        LinkModel model = new LinkModel
        {
            Url = "https://www.openai.com",
            LinkText = "OpenAI Homepage"
        };

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(resultPath);
    }
}
