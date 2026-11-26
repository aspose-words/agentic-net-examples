using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
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
        // Register code page provider required by Aspose.Words for certain encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON containing a Base64‑encoded attachment.
        string jsonData = @"
        {
            ""Title"": ""Sample Report"",
            ""HtmlSnippet"": ""<p>This is a <b>sample</b> HTML snippet embedded in the report.</p>"",
            ""AttachmentName"": ""SampleAttachment.txt"",
            ""AttachmentBase64"": ""U2FtcGxlIGF0dGFjaG1lbnQgY29udGVudC4gVGhpcyBpcyBhIHRlc3QgZmlsZS4=""
        }";

        // Deserialize JSON into the model.
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(jsonData)!;

        // Decode the attachment and write it to a file.
        byte[] attachmentBytes = Convert.FromBase64String(model.AttachmentBase64);
        File.WriteAllBytes(model.AttachmentName, attachmentBytes);
        model.AttachmentPath = Path.GetFullPath(model.AttachmentName);

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<<html [model.HtmlSnippet]>>");
        builder.Writeln($"Attachment: <<link [model.AttachmentPath] [model.AttachmentName]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
