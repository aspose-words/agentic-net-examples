using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph style to Heading 1 and write the heading text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

        // Prepare a simple (empty) data source for the ReportingEngine.
        // In this example the template does not contain any tags, so an empty anonymous object suffices.
        var dataSource = new { };

        // Initialize the ReportingEngine and populate the document.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource);

        // Save the resulting document to disk.
        doc.Save("LINQReporting.docx");
    }
}
