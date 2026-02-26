using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data source class that holds a document to be inserted.
    public class DocumentTestClass
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the RTF template that contains a <<doc [src.Document]>> tag.
            Document template = new Document("Template.rtf");

            // Load the document that will be inserted into the template.
            Document sourceDoc = new Document("Source.docx");

            // Prepare the data source for the reporting engine.
            var dataSource = new DocumentTestClass { Document = sourceDoc };

            // Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "src" matches the tag in the template.
            engine.BuildReport(template, dataSource, "src");

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
