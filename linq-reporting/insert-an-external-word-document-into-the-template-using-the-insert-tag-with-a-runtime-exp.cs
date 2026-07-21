using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create an external Word document that will be inserted.
        var externalDocPath = "external.docx";
        var externalDoc = new Document();
        var externalBuilder = new DocumentBuilder(externalDoc);
        externalBuilder.Writeln("This is content from the external document.");
        externalDoc.Save(externalDocPath);

        // Create the template document containing the <<doc>> tag.
        var templateDoc = new Document();
        var templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("=== Report Start ===");
        // Insert tag that references a runtime expression (src.Document).
        templateBuilder.Writeln("<<doc [src.Document]>>");
        templateBuilder.Writeln("=== Report End ===");

        // Prepare the data source wrapper exposing the external document.
        var src = new DocumentWrapper(new Document(externalDocPath));

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(templateDoc, src, "src");

        // Save the generated report.
        var outputPath = "result.docx";
        templateDoc.Save(outputPath);
    }
}

// Wrapper class exposing the external document as a public property.
public class DocumentWrapper
{
    public DocumentWrapper(Document document) => Document = document;

    public Document Document { get; set; }
}
