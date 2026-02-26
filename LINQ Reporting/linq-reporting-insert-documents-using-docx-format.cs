using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data source class that holds the document to be inserted.
    public class SourceDocument
    {
        public Document Document { get; set; }

        public SourceDocument(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a template document that contains a LINQ Reporting placeholder.
            // -------------------------------------------------
            Document template = new Document();                     // create blank document
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("=== Report Header ===");
            // The placeholder <<doc [src.Document]>> tells the ReportingEngine to insert the Document
            // stored in the data source object named "src".
            builder.Writeln("<<doc [src.Document]>>");
            builder.Writeln("=== Report Footer ===");

            // -------------------------------------------------
            // 2. Load the DOCX document that we want to embed into the template.
            // -------------------------------------------------
            Document docToInsert = new Document("Insert.docx");    // load existing DOCX file

            // -------------------------------------------------
            // 3. Wrap the loaded document in a data source object.
            // -------------------------------------------------
            var src = new SourceDocument(docToInsert);

            // -------------------------------------------------
            // 4. Build the final report using the LINQ Reporting engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The data source array contains the object, and the corresponding name array
            // provides the name ("src") used in the template placeholder.
            engine.BuildReport(template, new object[] { src }, new string[] { "src" });

            // -------------------------------------------------
            // 5. Save the resulting document.
            // -------------------------------------------------
            template.Save("Result.docx"); // save to file, format inferred from extension
        }
    }
}
