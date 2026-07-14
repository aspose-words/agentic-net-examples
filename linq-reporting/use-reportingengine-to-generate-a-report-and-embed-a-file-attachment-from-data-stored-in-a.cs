using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportModel
{
    public string Title { get; set; } = "";
    public string Content { get; set; } = "";
    public string AttachmentPath { get; set; } = "";
    public string AttachmentName { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string jsonPath = Path.Combine(workDir, "data.json");
        string attachmentPath = Path.Combine(workDir, "sample.txt");
        string outputPath = Path.Combine(workDir, "Report.docx");

        // Create a simple text file that will be attached.
        File.WriteAllText(attachmentPath, "This is the content of the attached file.");

        // Create JSON data that references the attachment.
        var model = new ReportModel
        {
            Title = "Report with Attachment",
            Content = "The following link points to an attached file.",
            AttachmentPath = attachmentPath,
            AttachmentName = "sample.txt"
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(model));

        // Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Content
        builder.Writeln("<<[model.Content]>>");
        builder.Writeln();

        // Attachment link (using the link tag)
        builder.Writeln("<<link [model.AttachmentPath] [model.AttachmentName]>>");
        builder.Writeln();

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Load JSON data source.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonData, "model");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}
