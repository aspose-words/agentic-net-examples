using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportModel
{
    public string CustomerName { get; set; } = "John Doe";
    public string ReportDate { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
    public string AttachmentFileName { get; set; } = "SampleAttachment.txt";
    public string AttachmentBase64 { get; set; } = "";
    public byte[] AttachmentContent => Convert.FromBase64String(AttachmentBase64);
}

public class Program
{
    public static void Main()
    {
        // Prepare sample attachment content and encode it as Base64.
        string attachmentText = "This is a sample attachment file.";
        byte[] attachmentBytes = Encoding.UTF8.GetBytes(attachmentText);
        string attachmentBase64 = Convert.ToBase64String(attachmentBytes);

        // Create the data model and fill the Base64 field.
        ReportModel model = new()
        {
            AttachmentBase64 = attachmentBase64
        };

        // Serialize the model to a JSON file (simulating an external source).
        string jsonPath = "data.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(model, Formatting.Indented));

        // Load the JSON back into a model instance.
        string jsonContent = File.ReadAllText(jsonPath);
        ReportModel data = JsonConvert.DeserializeObject<ReportModel>(jsonContent)!;

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new();
        DocumentBuilder builder = new(template);

        // Insert LINQ Reporting tags.
        builder.Writeln("Report for: <<[model.CustomerName]>>");
        builder.Writeln("Date: <<[model.ReportDate]>>");
        builder.Writeln("Attachment:");
        // Bookmark where the attachment will be inserted.
        builder.StartBookmark("Attachment");
        builder.Writeln("[Attachment will be inserted here]");
        builder.EndBookmark("Attachment");

        // Save the template to disk.
        string templatePath = "template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, data, "model");

        // -----------------------------------------------------------------
        // Embed the attachment as an OLE object at the bookmark location.
        // -----------------------------------------------------------------
        using (MemoryStream attachmentStream = new(data.AttachmentContent))
        {
            DocumentBuilder reportBuilder = new(reportDoc);
            reportBuilder.MoveToBookmark("Attachment");
            // Provide a valid ProgID for the OLE object (e.g., "Package" for generic files).
            reportBuilder.InsertOleObject(attachmentStream, "Package", false, null);
        }

        // Save the final report.
        string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
