using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsMhtmlInsertExample
{
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains MHTML files.
            string mhtmlFolder = @"C:\InputMhtml";

            // Load all MHTML documents from the folder using LINQ.
            // Each file is loaded into an Aspose.Words.Document instance.
            Document[] mhtmlDocuments = Directory
                .EnumerateFiles(mhtmlFolder, "*.mht")
                .Select(filePath => new Document(filePath)) // Load document from file.
                .ToArray();

            // Create a simple template document that will receive the MHTML contents.
            // The template uses the Reporting Engine syntax to insert a document.
            // <<doc [src.Document]>> will be replaced with the first document,
            // <<doc [src.Document] -sourceNumbering>> will insert the second, etc.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("=== Inserted MHTML Documents ===");
            // Insert placeholders for each document we plan to merge.
            for (int i = 0; i < mhtmlDocuments.Length; i++)
            {
                // The first placeholder uses the default insertion.
                // Subsequent placeholders use the -sourceNumbering option to keep original numbering.
                string placeholder = i == 0
                    ? "<<doc [src.Document]>>"
                    : "<<doc [src.Document] -sourceNumbering>>";
                builder.Writeln(placeholder);
            }

            // Use the ReportingEngine to populate the template.
            // The data source name "src" is referenced in the template placeholders.
            ReportingEngine engine = new ReportingEngine();
            // BuildReport expects an array of data sources and matching names.
            // Here we pass a single data source (the array of MHTML documents) and its name.
            engine.BuildReport(template, new object[] { mhtmlDocuments }, new string[] { "src" });

            // Save the resulting document.
            string outputPath = @"C:\Output\CombinedFromMhtml.docx";
            template.Save(outputPath);
        }
    }
}
