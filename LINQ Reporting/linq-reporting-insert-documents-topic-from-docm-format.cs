using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains the <<doc [src.Document]>> tag.
        Document template = new Document("Template.docm");

        // Load the document that will be inserted into the template.
        Document source = new Document("Source.docx");

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report, passing the source document as a data source named "src".
        // The template tag <<doc [src.Document]>> will be replaced with the contents of the source document.
        engine.BuildReport(template, new object[] { source }, new[] { "src" });

        // Save the merged result.
        template.Save("Result.docx");
    }
}
