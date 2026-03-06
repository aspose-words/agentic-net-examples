using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Simple class that holds a document to be inserted.
    public class DocumentSource
    {
        public Document Document { get; set; }

        public DocumentSource(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template (DOCM) that contains <<doc [src.Document]>> tags.
            Document template = new Document("Template.docm");

            // Prepare a list of source documents to be inserted.
            List<DocumentSource> sources = new List<DocumentSource>
            {
                new DocumentSource(new Document("Source1.docx")),
                new DocumentSource(new Document("Source2.docx"))
            };

            // Convert the list to the arrays required by ReportingEngine.
            object[] dataSources = new object[sources.Count];
            string[] dataSourceNames = new string[sources.Count];

            for (int i = 0; i < sources.Count; i++)
            {
                dataSources[i] = sources[i];
                dataSourceNames[i] = "src"; // All tags use the same name in the template.
            }

            // Build the report – this will replace the <<doc [src.Document]>> tags
            // with the contents of the corresponding source documents.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSources, dataSourceNames);

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
