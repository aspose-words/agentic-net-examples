using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark tag whose name comes from the model's BookmarkName property.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("This text is inside the bookmark.");
        builder.Writeln("<</bookmark>>");

        // Prepare the data model. The BookmarkName property must be a non‑empty string.
        ReportModel model = new()
        {
            BookmarkName = "MyBookmark"
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("BookmarkReport.docx");
    }
}

// Simple data model with a public property used in the bookmark expression.
public class ReportModel
{
    // Initialized to a non‑empty value to satisfy the bookmark requirement.
    public string BookmarkName { get; set; } = "DefaultBookmark";
}
