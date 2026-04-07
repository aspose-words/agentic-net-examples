using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // The document that will be inserted into the template.
    public Document Document { get; set; } = new Document();
}

public class Program
{
    public static void Main()
    {
        // ---------- Create the external document ----------
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the content of the external document.");
        const string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Load the external document so it can be passed to the reporting engine.
        var externalDoc = new Document(sourcePath);

        // ---------- Create the template document ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Define a bookmark where the external document will be inserted.
        builder.StartBookmark("InsertHere");

        // LINQ Reporting tag that includes another document.
        // The tag references the data source named "src" and its Document property.
        builder.Writeln("<<doc [src.Document]>>");

        builder.EndBookmark("InsertHere");

        // ---------- Prepare the data model ----------
        var model = new Model { Document = externalDoc };

        // ---------- Build the report ----------
        var engine = new ReportingEngine();
        // The data source name used in the tag is "src".
        engine.BuildReport(templateDoc, model, "src");

        // ---------- Save the result ----------
        const string outputPath = "Result.docx";
        templateDoc.Save(outputPath);
    }
}
