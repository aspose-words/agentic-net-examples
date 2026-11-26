using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // No data needed for this example.
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths used in the example.
        const string externalDocPath = "External.docx";
        const string templateDocPath = "Template.docx";
        const string resultDocPath = "Result.docx";

        // -----------------------------------------------------------------
        // Create an external DOCX file that will be included in the report.
        // -----------------------------------------------------------------
        var externalDoc = new Document();
        var externalBuilder = new DocumentBuilder(externalDoc);
        externalBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalDocPath);

        // ---------------------------------------------------------------
        // Create the template document with a bookmark that contains a <<doc>> tag.
        // ---------------------------------------------------------------
        var templateDoc = new Document();
        var templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Report start");
        // Bookmark opening tag with literal name "InsertHere".
        templateBuilder.Writeln("<<bookmark [\"InsertHere\"]>>");
        // Insert external document at the bookmark location using the supported <<doc>> tag.
        templateBuilder.Writeln($"<<doc [\"{externalDocPath}\"]>>");
        // Bookmark closing tag.
        templateBuilder.Writeln("<</bookmark>>");
        templateBuilder.Writeln("Report end");
        templateDoc.Save(templateDocPath);

        // ---------------------------------------------------------------
        // Load the template and build the report using LINQ Reporting Engine.
        // ---------------------------------------------------------------
        var loadedTemplate = new Document(templateDocPath);
        var engine = new ReportingEngine();

        // No data is required; an empty model is sufficient.
        engine.BuildReport(loadedTemplate, new ReportModel(), "model");

        // Save the final document.
        loadedTemplate.Save(resultDocPath);
    }
}
