using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple wrapper class used as a data source for the ReportingEngine.
    public class DocumentTestClass
    {
        public Document Document { get; set; }

        public DocumentTestClass(Document document)
        {
            Document = document;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create an HTML template that contains the ReportingEngine tags.
            //    The tags <<doc [src.Document]>> and <<doc [src.Document] -sourceNumbering>> will be
            //    replaced with the content of the source document at runtime.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            string htmlTemplate =
                "<<doc [src.Document]>>" + Environment.NewLine +
                "<<doc [src.Document] -sourceNumbering>>";

            builder.InsertHtml(htmlTemplate);

            // 2. Load the document that will be inserted dynamically.
            //    Replace the path with the actual location of your source document.
            Document sourceDoc = new Document("SourceDocument.docx");

            // 3. Prepare the data source for the ReportingEngine.
            //    The engine expects an array of objects and a matching array of names.
            var dataSource = new object[] { new DocumentTestClass(sourceDoc) };
            var dataSourceNames = new string[] { "src" };

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(template, dataSource, dataSourceNames);

            // 5. Save the resulting document.
            template.Save("ResultDocument.docx");
        }
    }
}
