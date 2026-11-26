using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Path to the external Word document to be inserted.
    public string ExternalDocPath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file names.
        string externalDocPath = "external.docx";
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create the external document that will be inserted into the report.
        // -----------------------------------------------------------------
        var externalDoc = new Document();
        var externalBuilder = new DocumentBuilder(externalDoc);
        externalBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalDocPath);

        // ---------------------------------------------------------------
        // Create the template document containing the <<doc>> insert tag.
        // ---------------------------------------------------------------
        var templateDoc = new Document();
        var templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Template start.");
        // Insert tag with a runtime expression that resolves to the external document path.
        templateBuilder.Writeln("<<doc [model.ExternalDocPath]>>");
        templateBuilder.Writeln("Template end.");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        var loadedTemplate = new Document(templatePath);

        // -------------------------------------------------
        // Prepare the root data model.
        // -------------------------------------------------
        var model = new ReportModel
        {
            ExternalDocPath = externalDocPath
        };

        // -------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // -------------------------------------------------
        // Save the final document.
        // -------------------------------------------------
        loadedTemplate.Save(outputPath);
    }
}
