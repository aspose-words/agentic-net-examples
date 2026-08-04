using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportData
{
    // The external document to be inserted.
    public Document Document { get; set; }

    public ReportData()
    {
        // Initialize to avoid nullable warnings.
        Document = new Document();
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the files used in the example.
        const string externalDocPath = "ExternalDocument.docx";
        const string templatePath = "Template.docx";
        const string outputPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create an external Word document that will be inserted later.
        // -----------------------------------------------------------------
        Document externalDoc = new Document();
        DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
        extBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalDocPath);

        // ---------------------------------------------------------------
        // 2. Create a template document containing the <<doc>> tag.
        // ---------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Report start");
        // The tag inserts the document referenced by src.Document at runtime.
        tmplBuilder.Writeln("<<doc [src.Document]>>");
        tmplBuilder.Writeln("Report end");
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template back from disk (required before building).
        // ---------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // ---------------------------------------------------------------
        // 4. Prepare the data source with the external document.
        // ---------------------------------------------------------------
        ReportData data = new ReportData
        {
            Document = new Document(externalDocPath)
        };

        // ---------------------------------------------------------------
        // 5. Build the report using the LINQ Reporting engine.
        // ---------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options are needed for this simple scenario.
        engine.BuildReport(loadedTemplate, data, "src");

        // ---------------------------------------------------------------
        // 6. Save the final document.
        // ---------------------------------------------------------------
        loadedTemplate.Save(outputPath);
    }
}
