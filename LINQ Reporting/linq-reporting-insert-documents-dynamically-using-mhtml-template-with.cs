using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data holder for a document that will be inserted.
    public class DocumentSource
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the MHTML template that contains the build switch tags.
            //    Example tag in the template: <<doc [src.Document]>> or <<doc [src.Document] -sourceNumbering>>
            Document template = new Document("Template.mht");

            // 2. Prepare the documents that need to be inserted dynamically.
            var docsToInsert = new List<DocumentSource>
            {
                new DocumentSource { Document = new Document("Insert1.docx") },
                new DocumentSource { Document = new Document("Insert2.docx") },
                new DocumentSource { Document = new Document("Insert3.docx") }
            };

            // 3. Build the data source arrays required by ReportingEngine.BuildReport.
            //    The engine expects an array of objects (our DocumentSource instances) and a matching
            //    array of names that will be used inside the template to reference each source.
            object[] dataSources = new object[docsToInsert.Count];
            string[] dataSourceNames = new string[docsToInsert.Count];

            for (int i = 0; i < docsToInsert.Count; i++)
            {
                dataSources[i] = docsToInsert[i];
                // The name must match the identifier used in the template (e.g., "src").
                // If the template uses the same name for all inserts, we can reuse it.
                // Here we use "src" for every source because the template tag accesses the
                // Document property of the object, not the object itself.
                dataSourceNames[i] = "src";
            }

            // 4. Create and configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after inserting documents.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // 5. Build the report – this will replace the <<doc [src.Document]>> tags
            //    with the actual content of each document in the dataSources array.
            engine.BuildReport(template, dataSources, dataSourceNames);

            // 6. Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
