using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqInsert
{
    // Simple wrapper that exposes a Document property for the reporting engine.
    public class DocumentWrapper
    {
        public Document Document { get; set; }

        public DocumentWrapper(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder that contains WORDML (*.xml) source documents.
            string wordmlFolder = @"C:\WordmlSources";

            // Load all WORDML files from the folder using LINQ.
            DocumentWrapper[] sourceDocs = Directory.GetFiles(wordmlFolder, "*.xml")
                .Select(path => new DocumentWrapper(new Document(path))) // Load each WORDML as a Document.
                .ToArray();

            // Create a template document on the fly.
            // The template contains a reporting tag that will be replaced with each source document.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("=== Begin Inserted Documents ===");
            // The tag <<doc [src.Document]>> tells the ReportingEngine to insert the Document property
            // of the data source named "src" at this location.
            builder.Writeln("<<doc [src.Document]>>");
            builder.Writeln("=== End Inserted Documents ===");

            // Build the report: insert each source document into the template.
            ReportingEngine engine = new ReportingEngine();
            // The data source name must match the name used in the tag ("src").
            engine.BuildReport(template, sourceDocs, new[] { "src" });

            // Save the resulting document.
            string outputPath = @"C:\Result\Combined.docx";
            template.Save(outputPath);
        }
    }
}
