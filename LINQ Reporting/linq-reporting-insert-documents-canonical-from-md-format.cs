using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // 1. Create a template document that contains a placeholder for the inserted markdown document.
        Document template = new Document();                     // create blank document
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report generated on " + DateTime.Now);
        builder.Writeln("<<doc [src.Document]>>"); // placeholder that ReportingEngine will replace

        // 2. Load the source markdown file. Aspose.Words auto‑detects the .md format.
        Document markdownDoc = new Document("source.md"); // load markdown as a Word document

        // 3. Wrap the markdown document in a simple data‑source class.
        var dataSource = new { src = new MarkdownWrapper { Document = markdownDoc } };

        // 4. Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "src");

        // 5. Save the final document.
        template.Save("Result.docx");
    }

    // Helper class exposing the Document property required by the <<doc [src.Document]>> tag.
    public class MarkdownWrapper
    {
        public Document Document { get; set; }
        public MarkdownWrapper() { }
        public MarkdownWrapper(Document doc) => Document = doc;
    }
}
