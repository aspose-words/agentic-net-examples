using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace DynamicDocumentInsertion
{
    // Simple data source class that holds the document to be inserted.
    public class DocumentTestClass
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the template that contains the build switch tags, e.g.:
            // <<doc [src.Document]>> and <<doc [src.Document] -sourceNumbering>>
            Document template = new Document("Template.docx");

            // Load the document that will be inserted into the template.
            Document sourceDoc = new Document("Source.docx");

            // Prepare the data source object for the reporting engine.
            var dataSource = new DocumentTestClass { Document = sourceDoc };

            // Create the reporting engine and populate the template.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter is the name used in the template to reference the data source.
            engine.BuildReport(template, dataSource, "src");

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
