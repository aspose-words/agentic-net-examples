using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace DynamicDocumentInsertion
{
    // Simple data source class that holds a Document to be inserted.
    public class DocumentSource
    {
        public Document Document { get; set; }

        public DocumentSource(Document document)
        {
            Document = document;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTX template that contains the <<doc [src.Document]>> tags.
            Document template = new Document("Template.dotx");

            // Load the document that will be inserted into the template.
            Document documentToInsert = new Document("Insert.docx");

            // Create the data source object and assign the document to be inserted.
            DocumentSource src = new DocumentSource(documentToInsert);

            // Use ReportingEngine to process the template.
            // The data source name "src" matches the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, new object[] { src }, new string[] { "src" });

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
