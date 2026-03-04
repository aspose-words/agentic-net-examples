using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Simple wrapper class that holds a Document.
    // The ReportingEngine can reference the Document via the property name in the template.
    public class DocumentWrapper
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains the reporting tags, e.g. <<doc [src.Document]>>.
            Document template = new Document("Template.docm");

            // Gather all source DOCX files that we want to insert into the template.
            // Adjust the folder path as needed.
            string sourceFolder = "SourceDocs";
            List<DocumentWrapper> sourceDocs = Directory
                .EnumerateFiles(sourceFolder, "*.docx")
                .Select(filePath => new DocumentWrapper { Document = new Document(filePath) })
                .ToList();

            // Prepare the arrays required by ReportingEngine.BuildReport.
            // The template expects a data source named "src".
            object[] dataSources = sourceDocs.Cast<object>().ToArray();
            string[] dataSourceNames = new[] { "src" };

            // Build the report – the engine will replace the tags in the template
            // with the contents of each source document.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSources, dataSourceNames);

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
