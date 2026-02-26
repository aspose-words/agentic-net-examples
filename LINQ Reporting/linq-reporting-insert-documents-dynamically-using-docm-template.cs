using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains tags like <<doc [src.Document]>>.
        Document template = new Document("Template.docm");

        // Load the document that will be inserted into the template.
        Document subDocument = new Document("Insert.docx");

        // Wrap the sub‑document in a simple data‑source class.
        var source = new DocumentSource { Document = subDocument };

        // Create the reporting engine and populate the template.
        ReportingEngine engine = new ReportingEngine();

        // The name "src" must match the tag used in the template.
        engine.BuildReport(template, source, "src");

        // Save the final document.
        template.Save("Result.docx");
    }

    // Data‑source class exposing the document to be inserted.
    public class DocumentSource
    {
        public Document Document { get; set; }
    }
}
