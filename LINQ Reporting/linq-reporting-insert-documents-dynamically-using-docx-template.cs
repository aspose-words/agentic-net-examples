using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Wrapper class that holds a Document to be inserted via the template.
    public class DocumentSource
    {
        public Document Document { get; }

        public DocumentSource(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains LINQ Reporting tags,
            // e.g. <<foreach [src]>> <<doc [src.Document]>> <<endforeach>>
            Document template = new Document("Template.docx");

            // Prepare a collection of documents that will be inserted dynamically.
            List<DocumentSource> sources = new List<DocumentSource>
            {
                new DocumentSource(new Document("InsertPart1.docx")),
                new DocumentSource(new Document("InsertPart2.docx")),
                new DocumentSource(new Document("InsertPart3.docx"))
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the overload that accepts a data source object and a name.
            // The name "src" must match the name used in the template tags.
            engine.BuildReport(template, sources, "src");

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
