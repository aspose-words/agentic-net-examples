using Aspose.Words;
using Aspose.Words.Reporting;

class SourceDocument
{
    public Document Document { get; set; }
}

class Program
{
    static void Main()
    {
        // Load the RTF template that contains a <<doc [src.Document]>> tag (or with a build switch).
        Document template = new Document("Template.rtf");

        // Create the document that will be inserted dynamically.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the first paragraph of the inserted document.");
        insertBuilder.InsertBreak(BreakType.PageBreak);
        insertBuilder.Writeln("This is the second paragraph on a new page.");

        // Prepare the data source for the ReportingEngine.
        var dataSource = new SourceDocument { Document = insertDoc };

        // Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "src");

        // Save the final document.
        template.Save("Result.docx");
    }
}
