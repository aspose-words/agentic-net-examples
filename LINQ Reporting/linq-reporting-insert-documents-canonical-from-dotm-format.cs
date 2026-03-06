using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data container that will be used as a data source for the reporting engine.
    public class DocumentDataSource
    {
        // The property name must match the name used in the template (e.g. <<doc [src.Document]>>).
        public Document Document { get; set; }

        public DocumentDataSource(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTM template that contains a reporting tag like <<doc [src.Document]>>.
            string templatePath = @"C:\Templates\ReportTemplate.dotm";

            // Load the DOTM template.
            Document template = new Document(templatePath);

            // Assume we have a folder with several source documents (DOCX) that we want to insert.
            string sourceDocsFolder = @"C:\SourceDocs";

            // Use LINQ to pick the first three .docx files (or any custom logic you need).
            List<Document> sourceDocuments = Directory
                .EnumerateFiles(sourceDocsFolder, "*.docx")
                .Take(3) // example: take first three files
                .Select(file => new Document(file)) // load each file into a Document object
                .ToList();

            // For demonstration we will insert the first document from the list.
            // Wrap it in the data source class expected by the reporting engine.
            DocumentDataSource dataSource = new DocumentDataSource(sourceDocuments.First());

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The second parameter is the data source object,
            // the third parameter is the name used in the template to reference the source.
            engine.BuildReport(template, dataSource, "src");

            // Save the populated report to a new file.
            string outputPath = @"C:\Output\GeneratedReport.docx";
            template.Save(outputPath, SaveFormat.Docx);
        }
    }
}
