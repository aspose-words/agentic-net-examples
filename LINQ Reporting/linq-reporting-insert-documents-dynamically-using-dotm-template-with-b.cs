using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model that will be passed to the reporting engine.
    // The property name "Document" matches the name used in the template tag <<doc [src.Document]>>.
    public class DocumentData
    {
        public Document Document { get; }

        public DocumentData(Document document)
        {
            Document = document;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTM template that contains the build switch tag.
            // Example tag inside the template: <<doc [src.Document]>> or <<doc [src.Document] -sourceNumbering>>
            const string templatePath = @"C:\Templates\DynamicInsertTemplate.dotm";

            // Path to the document that will be inserted dynamically.
            const string insertDocPath = @"C:\Documents\InsertMe.docx";

            // Load the template (DOTM) and the document to be inserted.
            Document template = new Document(templatePath);
            Document docToInsert = new Document(insertDocPath);

            // Prepare the data source – an array with a single object that holds the document to insert.
            var dataSource = new object[] { new DocumentData(docToInsert) };
            var dataSourceNames = new[] { "src" }; // "src" matches the name used in the template tag.

            // Configure the reporting engine. RemoveEmptyParagraphs is optional but often useful.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report – the engine will replace the <<doc [src.Document]>> tag with the content of docToInsert.
            engine.BuildReport(template, dataSource, dataSourceNames);

            // Save the resulting document.
            const string outputPath = @"C:\Output\DynamicInsertResult.docx";
            template.Save(outputPath);
        }
    }
}
