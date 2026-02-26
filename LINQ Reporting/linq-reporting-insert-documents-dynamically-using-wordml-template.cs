using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Simple wrapper class that will be used as a data source for the reporting engine.
    // The reporting engine can access public properties of this class from the template.
    public class DocumentWrapper
    {
        public Document Document { get; }

        public DocumentWrapper(Document document)
        {
            Document = document;
        }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a WORDML (or DOCX) template that contains the LINQ
            //    Reporting tags for inserting another document.
            //    The syntax <<doc [src.Document]>> tells the engine to insert the
            //    document referenced by the data source named "src".
            // -----------------------------------------------------------------
            Document template = new Document();                     // create a blank document
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("=== Start of Template ===");
            // Insert the first tag – normal insertion (numbering will continue).
            builder.Writeln("<<doc [src.Document]>>");
            // Insert the second tag – source numbering will be kept as is.
            builder.Writeln("<<doc [src.Document] -sourceNumbering>>");
            builder.Writeln("=== End of Template ===");

            // -----------------------------------------------------------------
            // 2. Load the document that we want to insert dynamically.
            //    In a real scenario this could be any existing .docx file.
            // -----------------------------------------------------------------
            // For demonstration we create a simple document on the fly.
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
            srcBuilder.Writeln("This is the content of the inserted document.");
            srcBuilder.Writeln("It can contain multiple paragraphs, tables, etc.");

            // Wrap the source document in a class that the reporting engine can use.
            DocumentWrapper dataSource = new DocumentWrapper(sourceDoc);

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            //    We pass the template, an array with a single data source object,
            //    and an array with the corresponding name ("src").
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove empty paragraphs that may appear after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            engine.BuildReport(
                template,                                 // the template containing <<doc ...>> tags
                new object[] { dataSource },              // array of data sources
                new[] { "src" }                          // names used inside the template
            );

            // -----------------------------------------------------------------
            // 4. Save the resulting document.
            // -----------------------------------------------------------------
            template.Save("LinqReportingResult.docx", SaveFormat.Docx);
        }
    }
}
