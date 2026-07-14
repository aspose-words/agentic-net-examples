using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Wrapper class for the external document source used in the <<doc>> tag.
    public class Source
    {
        public Document Document { get; set; } = null!;
    }

    public static void Main()
    {
        // Ensure the Aspose.Words license is not required for this example.

        // 1. Create the external document that will be inserted.
        Document externalDoc = new Document();
        var extBuilder = new DocumentBuilder(externalDoc);
        extBuilder.Writeln("This is the content of the external document.");
        const string externalPath = "External.docx";
        externalDoc.Save(externalPath);

        // 2. Create the template document containing a bookmark with the <<doc>> tag.
        Document templateDoc = new Document();
        var tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Report start");
        tmplBuilder.StartBookmark("InsertHere");
        // The <<doc>> tag inserts the document referenced by src.Document at this position.
        tmplBuilder.Writeln("<<doc [src.Document]>>");
        tmplBuilder.EndBookmark("InsertHere");
        tmplBuilder.Writeln("Report end");
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // 4. Prepare the data source for the reporting engine.
        var source = new Source { Document = new Document(externalPath) };

        // 5. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, source, "src");

        // 6. Save the final document.
        const string outputPath = "Result.docx";
        loadedTemplate.Save(outputPath);
    }
}
