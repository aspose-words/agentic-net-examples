using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTM template that contains a <<doc [src.Document] -sourceStyles>> tag.
        Document template = new Document("Template.dotm");

        // Load the document that will be inserted dynamically.
        Document sourceDoc = new Document("Source.docx");

        // Wrap the source document in a simple holder class so the template can reference it.
        var data = new DocumentHolder { Document = sourceDoc };

        // Build the report. The -sourceStyles switch tells the engine to keep the source
        // document's styles when inserting it into the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, new object[] { data }, new[] { "src" });

        // Save the final document.
        template.Save("Result.docx");
    }

    // Data source class used by the template (exposes a Document property).
    public class DocumentHolder
    {
        public Document Document { get; set; }
    }
}
