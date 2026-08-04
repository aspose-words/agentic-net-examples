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

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Paragraph that will be removed if the condition is false.
        // The <<if>> tag evaluates the boolean property Include.
        // When Include is false the paragraph becomes empty.
        builder.Writeln("<<if [model.Include]>><<[model.Text]>> <</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Load the template for reporting ----------
        Document doc = new Document(templatePath);

        // Sample data model. Include is false, so the paragraph will be empty.
        ReportModel model = new ReportModel
        {
            Include = false,
            Text = "This text will appear only when Include is true."
        };

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report. The root object name must match the tag prefix (model).
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        doc.Save(reportPath);
    }
}

// Simple data model used by the template.
public class ReportModel
{
    // When true the paragraph will contain Text; otherwise it will be empty.
    public bool Include { get; set; } = false;

    // Text to display if Include is true.
    public string Text { get; set; } = string.Empty;
}
