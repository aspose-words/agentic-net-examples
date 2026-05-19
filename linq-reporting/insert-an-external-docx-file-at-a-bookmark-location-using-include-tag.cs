using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string includeDocPath = "Included.docx";
        const string templatePath = "Template.docx";
        const string outputPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create the external document that will be inserted.
        // -----------------------------------------------------------------
        Document includeDoc = new Document();
        DocumentBuilder incBuilder = new DocumentBuilder(includeDoc);
        incBuilder.Writeln("This paragraph comes from the external document.");
        includeDoc.Save(includeDocPath);

        // -----------------------------------------------------------------
        // 2. Create the template document with a bookmark and an <<doc>> tag.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);

        // Insert a bookmark that will contain the included document.
        tmplBuilder.Writeln("Before the bookmark.");
        tmplBuilder.Writeln("<<bookmark [model.BookmarkName]>>");
        tmplBuilder.Writeln("<<doc [model.IncludeDoc]>>");
        tmplBuilder.Writeln("<</bookmark>>");
        tmplBuilder.Writeln("After the bookmark.");

        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template (as required by the reporting workflow).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            BookmarkName = "InsertHere",
            IncludeDoc = new Document(includeDocPath)
        };

        // -----------------------------------------------------------------
        // 5. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the final document.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Name of the bookmark where the external document will be placed.
    public string BookmarkName { get; set; } = "InsertHere";

    // The external document to be included.
    public Document IncludeDoc { get; set; } = new Document();
}
