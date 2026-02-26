using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a blank template document.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // 2. Insert a LINQ Reporting tag that will be replaced by the source document.
            //    The tag syntax "<<doc [src.Document]>>" tells the engine to insert the
            //    document referenced by the data source named "src".
            builder.Writeln("<<doc [src.Document]>>");

            // 3. Load the document that we want to insert.
            //    Replace the path with the actual location of your source .docx file.
            Document sourceDoc = new Document("Source.docx");

            // 4. Prepare the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // 5. Build the report.
            //    Pass the source document as a data source and give it the name "src"
            //    so that the template tag can reference it.
            engine.BuildReport(
                template,                                 // template document
                new object[] { sourceDoc },               // array of data sources
                new string[] { "src" }                    // corresponding names
            );

            // 6. Save the populated document.
            template.Save("Result.docx");
        }
    }
}
