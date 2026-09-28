using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // File names used in the example.
        const string externalDocPath = "external.docx";
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create the external Word document that will be inserted later.
        // -----------------------------------------------------------------
        Document externalDoc = new Document();
        DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
        extBuilder.Writeln("This is content from the external document.");
        externalDoc.Save(externalDocPath);

        // -----------------------------------------------------------------
        // Create the template document containing the <<doc>> tag.
        // The tag uses a runtime expression that returns the path to the
        // external document (model.ExternalDocPath).
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("<<doc [model.ExternalDocPath]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new()
        {
            ExternalDocPath = externalDocPath
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        reportDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path to the external document to be inserted.
    public string ExternalDocPath { get; set; } = string.Empty;
}
