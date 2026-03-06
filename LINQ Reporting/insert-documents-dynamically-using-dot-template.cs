using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOT template that contains the <<doc [src.Document]>> tag.
        Document template = new Document("Template.dotx");

        // Create a source document that will be inserted into the template.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the dynamically inserted content.");

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Populate the template with the source document.
        // The name "src" matches the tag in the template (<<doc [src.Document]>>).
        engine.BuildReport(template, new object[] { sourceDoc }, new string[] { "src" });

        // Save the final document.
        template.Save("Result.docx");
    }
}
