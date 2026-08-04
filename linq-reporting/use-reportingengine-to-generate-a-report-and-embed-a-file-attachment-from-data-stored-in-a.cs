using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string AttachmentPath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings used by Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file names.
        const string outputFolder = "Output";
        const string templatePath = "Output/template.docx";
        const string jsonPath = "Output/data.json";
        const string attachmentPath = "Output/attachment.txt";
        const string resultPath = "Output/ReportResult.docx";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create a simple attachment file that will be embedded.
        // -----------------------------------------------------------------
        File.WriteAllText(attachmentPath, "This is the content of the embedded attachment.", Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Create a JSON file that holds the data for the report.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Title = "Sample LINQ Reporting",
            AttachmentPath = attachmentPath // Path to the file we just created.
        };
        string jsonContent = JsonConvert.SerializeObject(model, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 3. Build the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a title placeholder.
        builder.Writeln("Report Title: <<[model.Title]>>");
        builder.Writeln();

        // Insert the document (attachment) placeholder.
        // The <<doc>> tag embeds the document referenced by the expression.
        builder.Writeln("Embedded Attachment:");
        builder.Writeln("<<doc [model.AttachmentPath]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template and the JSON data source.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var jsonDataSource = new JsonDataSource(jsonPath);

        // -----------------------------------------------------------------
        // 5. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        // The root object name used in the template tags is "model".
        engine.BuildReport(loadedTemplate, jsonDataSource, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(resultPath);
    }
}
