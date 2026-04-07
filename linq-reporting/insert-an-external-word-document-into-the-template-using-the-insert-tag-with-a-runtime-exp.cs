using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingInsertDoc
{
    // Wrapper class that will be passed to the reporting engine.
    public class Source
    {
        // The external document to be inserted.
        public Document Document { get; set; }

        public Source(Document document)
        {
            Document = document;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the files used in the example.
            const string externalDocPath = "external.docx";
            const string templatePath = "template.docx";
            const string outputPath = "output.docx";

            // -----------------------------------------------------------------
            // 1. Create the external document that will be inserted later.
            // -----------------------------------------------------------------
            Document externalDoc = new Document();
            DocumentBuilder extBuilder = new DocumentBuilder(externalDoc);
            extBuilder.Writeln("This is the content of the external document.");
            externalDoc.Save(externalDocPath);

            // -----------------------------------------------------------------
            // 2. Create the template document containing the <<doc>> tag.
            //    The tag uses a runtime expression that refers to src.Document.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
            tmplBuilder.Writeln("=== Begin of Template ===");
            // Insert tag that will be replaced by the external document at runtime.
            tmplBuilder.Writeln("<<doc [src.Document]>>");
            tmplBuilder.Writeln("=== End of Template ===");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template back from disk (required before building the report).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 4. Prepare the data source – an instance of the wrapper class.
            // -----------------------------------------------------------------
            Source src = new Source(new Document(externalDocPath));

            // -----------------------------------------------------------------
            // 5. Build the report using the LINQ Reporting engine.
            //    The root object name must match the name used in the tag (src).
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, src, "src");

            // -----------------------------------------------------------------
            // 6. Save the final document.
            // -----------------------------------------------------------------
            loadedTemplate.Save(outputPath);
        }
    }
}
