using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple class that holds a document to be inserted.
    // The property name "Document" matches the name used in the template tag <<doc [src.Document]>>.
    public class DocumentTestClass
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the WORDML (DOCX) template that contains the build switch tag.
            // Example tag inside the template: <<doc [src.Document]>> or <<doc [src.Document] -sourceNumbering>>
            const string templatePath = @"Templates\ReportTemplate.docx";

            // Load the template document.
            Document template = new Document(templatePath);

            // Prepare a list of source documents that will be inserted dynamically.
            // In a real scenario these could be generated or loaded from various locations.
            List<DocumentTestClass> sources = new List<DocumentTestClass>
            {
                new DocumentTestClass { Document = new Document(@"Sources\Doc1.docx") },
                new DocumentTestClass { Document = new Document(@"Sources\Doc2.docx") },
                new DocumentTestClass { Document = new Document(@"Sources\Doc3.docx") }
            };

            // The ReportingEngine can work with multiple data sources.
            // We pass the array of source objects and an array with a single name ("src").
            // The name "src" is referenced in the template tags.
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove empty paragraphs that may appear after insertion.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report by populating the template with the source documents.
            // The overload that accepts object[] and string[] allows us to reference the data source by name.
            engine.BuildReport(template, new object[] { sources }, new[] { "src" });

            // Save the resulting document.
            const string outputPath = @"Output\GeneratedReport.docx";
            template.Save(outputPath);
        }
    }
}
