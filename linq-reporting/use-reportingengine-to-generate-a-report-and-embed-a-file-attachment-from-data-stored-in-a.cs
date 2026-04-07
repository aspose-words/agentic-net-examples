using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportModel
{
    // Title of the report.
    public string Title { get; set; } = string.Empty;

    // Path to the attachment file (relative to the executable directory).
    public string AttachmentPath { get; set; } = string.Empty;

    // Returns a Document object loaded from the attachment file.
    // The LINQ Reporting Engine can embed this document using the <<doc>> tag.
    public Document AttachmentDocument
    {
        get
        {
            // Guard against missing file.
            if (string.IsNullOrEmpty(AttachmentPath) || !File.Exists(AttachmentPath))
                return new Document(); // Empty document as fallback.

            return new Document(AttachmentPath);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample data files (attachment and JSON source).
        // -----------------------------------------------------------------
        string baseDir = Directory.GetCurrentDirectory();

        // Sample attachment file.
        string attachmentFile = Path.Combine(baseDir, "sample.txt");
        File.WriteAllText(attachmentFile, "This is the content of the attached file.");

        // Sample JSON file that references the attachment.
        string jsonFile = Path.Combine(baseDir, "data.json");
        var jsonContent = new
        {
            Title = "Demo Report",
            AttachmentPath = attachmentFile
        };
        File.WriteAllText(jsonFile, JsonConvert.SerializeObject(jsonContent, Formatting.Indented));

        // -----------------------------------------------------------------
        // 2. Deserialize JSON into a strongly‑typed model.
        // -----------------------------------------------------------------
        ReportModel model = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonFile))!;

        // -----------------------------------------------------------------
        // 3. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(baseDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a title placeholder.
        builder.Writeln("Report Title: <<[model.Title]>>");
        builder.Writeln();

        // Insert the attachment placeholder using the <<doc>> tag.
        // The tag expects a Document object, which we expose via model.AttachmentDocument.
        builder.Writeln("Embedded Attachment:");
        builder.Writeln("<<doc [model.AttachmentDocument]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this scenario.
        engine.Options = ReportBuildOptions.None;

        // Build the report using the model as the root data source.
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(baseDir, "output.docx");
        reportDoc.Save(outputPath);
    }
}
