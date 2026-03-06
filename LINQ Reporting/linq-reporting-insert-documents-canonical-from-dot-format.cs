using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple class that holds a Document to be inserted via the reporting engine.
    public class DocumentContainer
    {
        public Document Document { get; set; }

        public DocumentContainer(string documentPath)
        {
            // Load the document that will be inserted into the template.
            Document = new Document(documentPath);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank template document.
            Document template = new Document();

            // 2. Build the template using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(template);
            // Insert a reporting tag that will be replaced by the document stored in the data source.
            // The tag syntax "<<doc [src.Document]>>" tells the engine to insert the whole document.
            // The second tag demonstrates the "-sourceNumbering" option (keeps source numbering as‑is).
            builder.Writeln("<<doc [src.Document]>>");
            builder.Writeln("<<doc [src.Document] -sourceNumbering>>");

            // 3. Prepare the data source that contains the document to be inserted.
            // Replace "List item.docx" with the path to your source document.
            DocumentContainer src = new DocumentContainer("List item.docx");

            // 4. Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "src" matches the name used in the template tags.
            engine.BuildReport(template, new object[] { src }, new string[] { "src" });

            // 5. Save the resulting document.
            template.Save("ReportWithInsertedDocs.docx");
        }
    }
}
