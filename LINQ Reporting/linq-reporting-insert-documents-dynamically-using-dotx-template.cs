using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model that the template will reference.
    // The template should contain a tag like <<doc [src.Document]>>.
    public class DocItem
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTX template that contains the <<doc [src.Document]>> tag.
            const string templatePath = @"C:\Templates\ReportTemplate.dotx";

            // Path where the final report will be saved.
            const string outputPath = @"C:\Reports\GeneratedReport.docx";

            // Load the DOTX template.
            Document reportTemplate = new Document(templatePath);

            // Prepare a collection of documents that will be inserted dynamically.
            List<DocItem> itemsToInsert = new List<DocItem>
            {
                new DocItem { Document = CreateSampleDocument("First inserted document.") },
                new DocItem { Document = CreateSampleDocument("Second inserted document.") },
                new DocItem { Document = CreateSampleDocument("Third inserted document.") }
            };

            // ReportingEngine will replace the <<doc [src.Document]>> tag with the documents
            // from the data source. The data source name ("src") must match the tag prefix.
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove any empty paragraphs that may appear after insertion.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the template and the data source array.
            // The first (and only) data source name is "src".
            engine.BuildReport(reportTemplate,
                               new object[] { itemsToInsert },
                               new[] { "src" });

            // Save the generated report.
            reportTemplate.Save(outputPath);
        }

        // Helper method to create a simple one‑page document with a single paragraph.
        private static Document CreateSampleDocument(string text)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln(text);
            return doc;
        }
    }
}
