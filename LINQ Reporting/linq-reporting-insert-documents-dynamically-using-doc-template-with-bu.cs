using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    class Program
    {
        static void Main()
        {
            // Load the template that contains the build switches:
            //   <<doc [src.Document]>>               – inserts the document and continues numbering.
            //   <<doc [src.Document] -sourceNumbering>> – inserts the document and keeps its own numbering.
            Document template = new Document("TemplateWithBuildSwitch.docx");

            // Load the document that will be inserted dynamically.
            Document sourceDoc = new Document("DocumentToInsert.docx");

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Optional: remove empty paragraphs that may appear after processing.
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report. The data source name "src" matches the tag in the template.
            // The source document is passed inside an object array because the overload expects
            // an array of data sources and an array of corresponding names.
            engine.BuildReport(template,
                new object[] { sourceDoc },
                new string[] { "src" });

            // Save the resulting document.
            template.Save("ResultWithInsertedDocument.docx");
        }
    }
}
