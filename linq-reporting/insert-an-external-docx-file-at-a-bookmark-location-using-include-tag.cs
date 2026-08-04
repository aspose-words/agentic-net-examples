using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // File names used in the example.
        const string externalDocPath = "External.docx";
        const string templatePath = "Template.docx";
        const string outputPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create the external document that will be inserted later.
        // -----------------------------------------------------------------
        Document externalDoc = new Document();
        DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
        extBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save(externalDocPath);

        // -----------------------------------------------------------------
        // 2. Create the template document with a bookmark and a <<doc>> tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);

        tmplBuilder.Writeln("Start of the main document.");

        // Bookmark where the external document will be inserted.
        tmplBuilder.StartBookmark("InsertHere");
        // The <<doc>> tag tells the LINQ Reporting engine to insert the document.
        tmplBuilder.Writeln("<<doc [src.Document]>>");
        tmplBuilder.EndBookmark("InsertHere");

        tmplBuilder.Writeln("End of the main document.");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template for reporting.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Prepare the data model for the report.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Document = new Document(externalDocPath) // Load the external document.
        };

        // -----------------------------------------------------------------
        // 5. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The data source name "src" must match the prefix used in the <<doc>> tag.
        engine.BuildReport(loadedTemplate, model, "src");

        // -----------------------------------------------------------------
        // 6. Save the final document.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
// The property name must match the expression used in the <<doc>> tag.
public class ReportModel
{
    public Document Document { get; set; } = null!;
}
