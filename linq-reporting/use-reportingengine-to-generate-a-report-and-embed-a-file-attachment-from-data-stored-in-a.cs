using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Model that matches the JSON structure.
    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public string AttachmentPath { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string attachmentPath = Path.Combine(workDir, "attachment.docx");
        string jsonPath = Path.Combine(workDir, "data.json");
        string outputPath = Path.Combine(workDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple attachment document (will be embedded later).
        // -----------------------------------------------------------------
        Document attachmentDoc = new Document();
        DocumentBuilder attBuilder = new DocumentBuilder(attachmentDoc);
        attBuilder.Writeln("This is the embedded attachment document.");
        attachmentDoc.Save(attachmentPath);

        // ---------------------------------------------------------------
        // 2. Create JSON data that references the attachment document.
        // ---------------------------------------------------------------
        var model = new ReportModel
        {
            Title = "Sample Report",
            Body = "The report body is generated from JSON data.",
            AttachmentPath = attachmentPath // absolute path to the attachment
        };
        string jsonContent = JsonConvert.SerializeObject(model, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonContent);

        // ---------------------------------------------------------------
        // 3. Build a template document containing LINQ Reporting tags.
        // ---------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Body
        builder.Writeln("<<[model.Body]>>");
        builder.Writeln();

        // Embed the attachment document using the <<doc>> tag.
        // The tag will be resolved by the ReportingEngine at build time.
        builder.Writeln("Embedded attachment:");
        builder.Writeln("<<doc [model.AttachmentPath]>>");

        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 4. Load the template and generate the final report.
        // ---------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        ReportingEngine engine = new ReportingEngine();
        // No special options required for this scenario.
        engine.BuildReport(loadedTemplate, jsonDataSource, "model");

        // ---------------------------------------------------------------
        // 5. Save the generated report.
        // ---------------------------------------------------------------
        loadedTemplate.Save(outputPath);

        // Indicate completion (no interactive prompts as required).
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
