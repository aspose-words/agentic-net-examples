using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Class that holds a document which will be inserted into the template.
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
            // Directory for generated files.
            string outputDir = "Output/";
            System.IO.Directory.CreateDirectory(outputDir);

            // 1. Create a source document that will be inserted.
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
            srcBuilder.Writeln("This is the first inserted document.");
            srcBuilder.Writeln("It contains a simple paragraph.");
            sourceDoc.Save(outputDir + "Source.docx"); // optional, just for inspection

            // 2. Create a DOCM template containing the <<doc>> tag.
            Document template = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(template);
            // Tag without switch – numbering continues from the host document.
            tmplBuilder.Writeln("<<doc [src.Document]>>");
            tmplBuilder.Writeln(); // blank line for readability
            // Tag with -sourceNumbering switch – numbering of the inserted doc is kept as is.
            tmplBuilder.Writeln("<<doc [src.Document] -sourceNumbering>>");
            template.Save(outputDir + "Template.docm");

            // 3. Prepare the data source for the ReportingEngine.
            // The engine will reference the property "Document" via the name "src".
            DocumentTestClass data = new DocumentTestClass(sourceDoc);
            object[] dataSources = new object[] { data };
            string[] dataSourceNames = new string[] { "src" };

            // 4. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove any empty paragraphs left after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(template, dataSources, dataSourceNames);

            // 5. Save the final document.
            template.Save(outputDir + "Result.docx");
        }
    }
}
