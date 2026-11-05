using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // Paths for the files used in the example.
        string externalDocPath = Path.Combine(workDir, "external.docx");
        string templatePath = Path.Combine(workDir, "template.docx");
        string resultPath = Path.Combine(workDir, "result.docx");

        // -----------------------------------------------------------------
        // Create the external Word document that will be inserted later.
        // -----------------------------------------------------------------
        Document externalDoc = new Document();
        DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
        extBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalDocPath);

        // -----------------------------------------------------------------
        // Create the template document containing a placeholder for the external doc.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("=== Report Start ===");
        // The <<doc>> tag tells the LINQ Reporting engine to insert a document.
        builder.Writeln("<<doc [model.ExternalDocPath]>>");
        builder.Writeln("=== Report End ===");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Prepare the data model. The property points to the external document.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            ExternalDocPath = externalDocPath
        };

        // -----------------------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Save the final document.
        // -----------------------------------------------------------------
        reportDoc.Save(resultPath);
    }
}

// Data model used by the LINQ Reporting template.
public class ReportModel
{
    // Path to the external Word document to be inserted.
    public string ExternalDocPath { get; set; } = string.Empty;
}
