using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportModel
{
    public string Title { get; set; } = "";
    public string HtmlSnippet { get; set; } = "";
    public string AttachmentName { get; set; } = "";
    public string AttachmentBase64 { get; set; } = "";
    public string AttachmentPath { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample attachment content and encode it to Base64.
        string attachmentContent = "This is the content of the attachment file.";
        string attachmentBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(attachmentContent));

        // Build a model instance and serialize it to JSON (simulating external JSON source).
        var initialModel = new ReportModel
        {
            Title = "Sample Report",
            HtmlSnippet = "<b>Bold Text</b> and <i>Italic Text</i>",
            AttachmentName = "SampleAttachment.txt",
            AttachmentBase64 = attachmentBase64
        };
        string json = JsonConvert.SerializeObject(initialModel);

        // Deserialize JSON back to model.
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(json)!;

        // Decode attachment and write it to a file.
        byte[] attachmentBytes = Convert.FromBase64String(model.AttachmentBase64);
        string attachmentPath = Path.Combine(outputDir, model.AttachmentName);
        File.WriteAllBytes(attachmentPath, attachmentBytes);
        model.AttachmentPath = attachmentPath; // Update model for linking.

        // Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<<[model.HtmlSnippet] -html>>");
        builder.Writeln("Attachment: <<link [model.AttachmentPath] [model.AttachmentName]>>");

        templateDoc.Save(templatePath);

        // Load the template and generate the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
