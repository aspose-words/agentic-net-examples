using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InsertDocumentsWithTemplate
{
    static void Main()
    {
        // 1. Create a DOTM template that contains tags for inserting a document.
        // The tag <<doc [src.Document]>> inserts the document with default numbering.
        // The tag <<doc [src.Document] -sourceNumbering>> inserts the document and keeps its own numbering.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // First insertion – default behavior (numbering continues).
        builder.Writeln("<<doc [src.Document]>>");
        builder.Writeln(); // add an empty line between the two insertions

        // Second insertion – keep source numbering.
        builder.Writeln("<<doc [src.Document] -sourceNumbering>>");

        // Save the template as a DOTM file.
        const string templatePath = "Template.dotm";
        template.Save(templatePath, SaveFormat.Dotm);

        // 2. Load the source document that will be inserted into the template.
        // This can be any Word document (DOCX, DOC, etc.).
        const string sourceDocPath = "Source.docx";
        Document sourceDoc = new Document(sourceDocPath);

        // 3. Load the template back (demonstrates the load rule).
        Document loadedTemplate = new Document(templatePath);

        // 4. Prepare the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            // Optional: remove empty paragraphs that may be created by the engine.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // 5. Build the report.
        // The data source array contains the document to be inserted.
        // The corresponding name ("src") is used inside the template tags.
        engine.BuildReport(loadedTemplate,
                           new object[] { sourceDoc },
                           new string[] { "src" });

        // 6. Save the final document.
        const string resultPath = "Result.docx";
        loadedTemplate.Save(resultPath, SaveFormat.Docx);
    }
}
