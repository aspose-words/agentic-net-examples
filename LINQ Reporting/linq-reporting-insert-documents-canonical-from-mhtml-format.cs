using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsMhtmlLinqReporting
{
    class Program
    {
        static void Main()
        {
            // Folder that contains the source MHTML files.
            string mhtmlFolder = @"C:\MhtmlSource";

            // Load each MHTML file into a Document object using the Document constructor.
            // The constructor automatically detects the MHTML format.
            Document[] sourceDocuments = Directory.GetFiles(mhtmlFolder, "*.mhtml")
                .Select(filePath => new Document(filePath))
                .ToArray();

            // -----------------------------------------------------------------
            // Create a simple template document that will receive the inserted
            // documents. The template uses the Reporting Engine tag <<doc [src.Document]>>.
            // -----------------------------------------------------------------
            Document template = new Document();                     // create blank document
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("=== Begin Inserted Documents ===");
            // The tag tells ReportingEngine to insert each document from the data source.
            builder.Writeln("<<doc [src.Document]>>");
            builder.Writeln("=== End Inserted Documents ===");

            // -----------------------------------------------------------------
            // Build the report. The data source is an array containing the MHTML
            // documents, referenced in the template by the name "src".
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The first (and only) data source is the array of Document objects.
            engine.BuildReport(template, new object[] { sourceDocuments }, new string[] { "src" });

            // -----------------------------------------------------------------
            // Save the resulting document.
            // -----------------------------------------------------------------
            string outputPath = @"C:\Result\CombinedFromMhtml.docx";
            template.Save(outputPath, SaveFormat.Docx);

            Console.WriteLine("Report generated successfully at: " + outputPath);
        }
    }
}
