using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple class that will be used as a data source for the template.
    // The template will reference the property "Document" via the build switch.
    public class DocumentData
    {
        public Document Document { get; set; }

        public DocumentData(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTX template that contains the build switch tag:
            // <<doc [src.Document]>>   – inserts the document referenced by src.Document
            // <<doc [src.Document] -sourceNumbering>> – optional switch to keep source numbering.
            string templatePath = @"C:\Templates\InsertDocsTemplate.dotx";

            // Load the template document.
            Document template = new Document(templatePath);

            // Prepare the documents that will be inserted dynamically.
            // In a real scenario these could be generated on‑the‑fly or loaded from a database.
            var docsToInsert = new List<DocumentData>
            {
                new DocumentData(new Document(@"C:\SourceDocs\First.docx")),
                new DocumentData(new Document(@"C:\SourceDocs\Second.docx")),
                new DocumentData(new Document(@"C:\SourceDocs\Third.docx"))
            };

            // The ReportingEngine can work with multiple data sources.
            // Here we use a single data source array containing the list above.
            // The name "src" must match the name used in the template tag.
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove empty paragraphs that may appear after insertion.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report – the engine will replace the <<doc ...>> tags with the actual documents.
            engine.BuildReport(template, new object[] { docsToInsert }, new[] { "src" });

            // Save the resulting document.
            string outputPath = @"C:\Output\CombinedReport.docx";
            template.Save(outputPath);
        }
    }
}
