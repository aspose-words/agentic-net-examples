using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains the reporting tag <<doc [src.Document]>>.
        Document template = new Document("Template.docm");

        // Load the document that we want to insert into the template.
        Document sourceDocument = new Document("Source.docx");

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report, passing the source document as a data source.
        // The data source name "src" matches the tag used in the template.
        engine.BuildReport(template, new object[] { sourceDocument }, new string[] { "src" });

        // Save the resulting document.
        template.Save("Result.docx");
    }
}
