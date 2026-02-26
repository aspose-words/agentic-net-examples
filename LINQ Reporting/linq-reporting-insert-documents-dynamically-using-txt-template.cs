using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the TXT template that contains the LINQ Reporting tag <<doc [src.Document]>>.
        Document template = new Document("Template.txt");

        // Load the document that will be inserted into the template.
        Document docToInsert = new Document("Insert.docx");

        // Wrap the document in a simple data source class so it can be referenced from the template.
        var source = new DocumentSource { Document = docToInsert };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The template tag references the data source name "src".
        engine.BuildReport(template, source, "src");

        // Save the generated document.
        template.Save("Result.docx");
    }

    // Simple class used as a data source for the template.
    public class DocumentSource
    {
        public Document Document { get; set; }
    }
}
