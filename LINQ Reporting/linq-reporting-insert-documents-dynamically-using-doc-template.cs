using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple wrapper class that holds a Document to be inserted.
    public class DocumentTestClass
    {
        public Document Document { get; set; }

        public DocumentTestClass(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template that contains LINQ Reporting tags, e.g. <<doc [src.Document]>>
            Document template = new Document("Template.docx");

            // Load the documents that will be inserted dynamically.
            Document insertDoc1 = new Document("Insert1.docx");
            Document insertDoc2 = new Document("Insert2.docx");

            // Wrap the documents in a class that matches the tag reference in the template.
            DocumentTestClass src1 = new DocumentTestClass(insertDoc1);
            DocumentTestClass src2 = new DocumentTestClass(insertDoc2);

            // Create the ReportingEngine and configure it (optional).
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after tag replacement.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The data source name "src" must match the tag prefix used in the template.
            engine.BuildReport(template,
                new object[] { src1, src2 },          // array of data sources
                new[] { "src" });                    // corresponding names

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
