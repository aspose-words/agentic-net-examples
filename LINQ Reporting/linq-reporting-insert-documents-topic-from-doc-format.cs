using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains <<doc [src.Document]>> tags.
            string templatePath = @"Templates\ReportTemplate.docx";

            // Load the template document.
            Document template = new Document(templatePath);

            // Directory that holds the source DOCX files to be inserted.
            string sourceDocsDir = @"SourceDocs";

            // Load each source document using LINQ and create an array of objects.
            Document[] sourceDocuments = Directory
                .EnumerateFiles(sourceDocsDir, "*.docx")
                .Select(filePath => new Document(filePath))
                .ToArray();

            // Prepare data source names that correspond to the tags in the template.
            // For example, the template may contain tags like <<doc [src0.Document]>>.
            string[] dataSourceNames = sourceDocuments
                .Select((doc, index) => $"src{index}")
                .ToArray();

            // The ReportingEngine can accept multiple data sources.
            ReportingEngine engine = new ReportingEngine();

            // Build the report – this will replace the <<doc ...>> tags with the contents of the source documents.
            engine.BuildReport(template, sourceDocuments.Cast<object>().ToArray(), dataSourceNames);

            // Save the resulting document.
            string outputPath = @"Results\CombinedReport.docx";
            template.Save(outputPath);
        }
    }
}
