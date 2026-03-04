using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsPdfInsertExample
{
    // Simple wrapper class that holds a Document instance.
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
            // Folder that contains source PDF files.
            string pdfFolder = @"C:\SourcePdfs";

            // -----------------------------------------------------------------
            // 1. Create a Word template that contains a reporting tag.
            //    The tag <<doc [src.Document]>> will be replaced with each
            //    document from the data source during the report build.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Inserted PDFs follow:");
            // Reporting tag – the data source name will be \"src\".
            builder.Writeln("<<doc [src.Document]>>");
            builder.Writeln("\n--- End of inserted PDF ---");

            // -----------------------------------------------------------------
            // 2. Load all PDF files from the folder and wrap them.
            // -----------------------------------------------------------------
            DocumentWrapper[] pdfWrappers = Directory.GetFiles(pdfFolder, "*.pdf")
                .Select(pdfPath => new DocumentWrapper(new Document(pdfPath)))
                .ToArray();

            // -----------------------------------------------------------------
            // 3. Build the report – the ReportingEngine will replace the tag
            //    with each PDF document. The data source name is \"src\".
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template,
                new object[] { pdfWrappers },               // data source array
                new[] { "src" });                           // corresponding names

            // -----------------------------------------------------------------
            // 4. Save the resulting document.
            // -----------------------------------------------------------------
            string outputPath = @"C:\Result\CombinedFromPdfs.docx";
            template.Save(outputPath);
        }
    }
}
