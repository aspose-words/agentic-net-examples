using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple wrapper class used as a data source for the reporting engine.
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
            // Path to the template that contains the <<doc [src.Document]>> tag.
            string templatePath = @"C:\Docs\Template.docx";

            // Folder that contains the documents to be inserted.
            string sourceFolder = @"C:\Docs\Sources";

            // Load the template document.
            Document template = new Document(templatePath);

            // Load all DOCX files from the source folder.
            List<DocumentSource> sources = new List<DocumentSource>();
            foreach (string file in Directory.GetFiles(sourceFolder, "*.docx"))
            {
                Document srcDoc = new Document(file);
                sources.Add(new DocumentSource(srcDoc));
            }

            // Prepare arrays for the ReportingEngine overload that accepts multiple data sources.
            object[] dataSources = new object[sources.Count];
            string[] dataSourceNames = new string[sources.Count];

            for (int i = 0; i < sources.Count; i++)
            {
                dataSources[i] = sources[i];
                // The name "src" matches the tag used in the template.
                dataSourceNames[i] = "src";
            }

            // Build the report – this will replace the <<doc [src.Document]>> tag with each source document.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSources, dataSourceNames);

            // Save the resulting document.
            string outputPath = @"C:\Docs\Result.docx";
            template.Save(outputPath);
        }
    }
}
