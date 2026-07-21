using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string jsonPath = Path.Combine(workDir, "data.json");
        string attachmentPath = Path.Combine(workDir, "attachment.txt");
        string outputPath = Path.Combine(workDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple attachment file that will be embedded in the report.
        // -----------------------------------------------------------------
        File.WriteAllText(attachmentPath, "This is the content of the attached file.");

        // -----------------------------------------------------------------
        // 2. Create JSON data containing a title and the path to the attachment.
        // -----------------------------------------------------------------
        var jsonData = new
        {
            Title = "Report with Embedded Attachment",
            AttachmentPath = attachmentPath
        };
        string jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
        File.WriteAllText(jsonPath, jsonString);

        // -----------------------------------------------------------------
        // 3. Build a template document programmatically.
        //    The template uses LINQ Reporting tags to insert the title and the attachment.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln(); // empty line
        builder.Writeln("Attachment:");
        // The <<doc>> tag inserts the content of the file referenced by the expression.
        builder.Writeln("<<doc [model.AttachmentPath]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template and the JSON data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // -----------------------------------------------------------------
        // 5. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(reportDoc, dataSource, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
