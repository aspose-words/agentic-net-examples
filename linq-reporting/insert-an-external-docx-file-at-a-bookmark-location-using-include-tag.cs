using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string externalDocPath = Path.Combine(workDir, "external.docx");
        string outputPath = Path.Combine(workDir, "result.docx");

        // Create the external document that will be inserted.
        CreateExternalDocument(externalDocPath);

        // Create the LINQ Reporting template containing a bookmark and a doc tag.
        CreateTemplateDocument(templatePath);

        // Load the template.
        Document template = new Document(templatePath);

        // Prepare the model.
        ReportModel model = new()
        {
            BookmarkName = "InsertHere",
            IncludePath = externalDocPath
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        template.Save(outputPath);
    }

    private static void CreateExternalDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the content of the external document.");
        doc.Save(path);
    }

    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory text.
        builder.Writeln("Report generated with Aspose.Words LINQ Reporting.");
        builder.Writeln();

        // Bookmark start.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        // Insert external document at the bookmark location using the supported <<doc>> tag.
        builder.Writeln("<<doc [model.IncludePath]>>");
        // Bookmark end.
        builder.Writeln("<</bookmark>>");

        doc.Save(path);
    }

    // Model class used by the LINQ Reporting engine.
    public class ReportModel
    {
        public string BookmarkName { get; set; } = string.Empty;
        public string IncludePath { get; set; } = string.Empty;
    }
}
