using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample external document that will be inserted later.
        // -----------------------------------------------------------------
        Document externalDoc = new Document();
        DocumentBuilder externalBuilder = new DocumentBuilder(externalDoc);
        externalBuilder.Writeln("This is the content of the external document.");
        externalDoc.Save("External.docx");

        // Load the external document so it can be passed to the reporting engine.
        Document includeDoc = new Document("External.docx");

        // -----------------------------------------------------------------
        // Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report start");

        // Open a bookmark whose name comes from the model.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");

        // Insert the external document inside the bookmark using the supported <<doc>> tag.
        builder.Writeln("<<doc [model.IncludeDoc]>>");

        // Close the bookmark.
        builder.Writeln("<</bookmark>>");

        builder.Writeln("Report end");

        // (Optional) Save the template for inspection.
        template.Save("Template.docx");

        // -----------------------------------------------------------------
        // Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel(includeDoc);

        // -----------------------------------------------------------------
        // Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the final document.
        template.Save("Result.docx");
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Name of the bookmark where the external document will be inserted.
    public string BookmarkName { get; set; } = "InsertHere";

    // The external document to include.
    public Document IncludeDoc { get; set; }

    public ReportModel(Document includeDoc)
    {
        IncludeDoc = includeDoc ?? throw new ArgumentNullException(nameof(includeDoc));
    }
}
