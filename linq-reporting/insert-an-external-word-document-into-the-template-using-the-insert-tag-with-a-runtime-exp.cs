using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // The external document to be inserted.
    public Document Document { get; set; } = null!;
}

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create the external document that will be inserted.
        string externalPath = Path.Combine(workDir, "External.docx");
        Document externalDoc = new Document();
        DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
        extBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalPath);

        // 2. Create the template document containing the <<doc>> tag.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("=== Report Start ===");
        // Insert tag that references the external document via a runtime expression.
        tmplBuilder.Writeln("<<doc [src.Document]>>");
        tmplBuilder.Writeln("=== Report End ===");
        templateDoc.Save(templatePath);

        // 3. Load the template and the external document.
        Document loadedTemplate = new Document(templatePath);
        Document loadedExternal = new Document(externalPath);

        // 4. Prepare the data model for the ReportingEngine.
        ReportModel model = new ReportModel
        {
            Document = loadedExternal
        };

        // 5. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "src", matching the tag expression.
        engine.BuildReport(loadedTemplate, model, "src");

        // 6. Save the final document.
        string resultPath = Path.Combine(workDir, "Result.docx");
        loadedTemplate.Save(resultPath);
    }
}
