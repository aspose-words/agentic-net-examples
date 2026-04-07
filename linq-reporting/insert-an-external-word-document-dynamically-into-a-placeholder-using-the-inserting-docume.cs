using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the external document that will be inserted.
        Document externalDoc = new Document();
        DocumentBuilder externalBuilder = new DocumentBuilder(externalDoc);
        externalBuilder.Writeln("This is the content of the external document.");
        // Optionally save the external document (not required for the insertion itself).
        externalDoc.Save("External.docx");

        // Create the template document containing the placeholder tag.
        Document template = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(template);
        templateBuilder.Writeln("Document before insertion:");
        // Placeholder tag that tells the ReportingEngine to insert the document from src.Document.
        templateBuilder.Writeln("<<doc [src.Document]>>");
        templateBuilder.Writeln("Document after insertion.");
        template.Save("Template.docx");

        // Load the template document (could also reuse the in‑memory instance).
        Document loadedTemplate = new Document("Template.docx");

        // Prepare the data source with a public property named Document.
        var src = new Src { Document = externalDoc };

        // Build the report – the placeholder will be replaced with the external document.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, src, "src");

        // Save the final document.
        loadedTemplate.Save("Result.docx");
    }

    // Wrapper class used as the root data source for the LINQ Reporting engine.
    public class Src
    {
        // The property name matches the expression used in the template tag.
        public Document Document { get; set; } = new Document();
    }
}
