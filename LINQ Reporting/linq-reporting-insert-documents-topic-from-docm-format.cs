using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple class that holds a document to be inserted via the <<doc>> tag in the template.
    public class DocumentHolder
    {
        public Document Document { get; set; }

        public DocumentHolder(string filePath)
        {
            Document = new Document(filePath);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the template (DOCM) and the documents to be inserted.
            string dataDir = @"C:\Data\";

            // Load the DOCM template that contains Reporting Engine tags such as <<doc [src.Document]>>.
            Document template = new Document(dataDir + "Template.docm");

            // Prepare the data source that will be referenced in the template.
            // In this example we have a simple anonymous object with a title and a list of items.
            var reportData = new
            {
                Title = "Quarterly Sales Report",
                Items = new List<dynamic>
                {
                    new { Product = "Laptop", Quantity = 120, Amount = 150000 },
                    new { Product = "Smartphone", Quantity = 340, Amount = 255000 },
                    new { Product = "Tablet", Quantity = 210, Amount = 105000 }
                }
            };

            // Create a ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the anonymous object as the data source.
            // The template can reference fields like <<[Title]>> and iterate over <<foreach [Items]>><<[Product]>> etc.
            engine.BuildReport(template, reportData);

            // -------------------------------------------------------------------------
            // Insert additional whole documents into the report using the <<doc>> tag.
            // -------------------------------------------------------------------------

            // Load the documents that will be inserted.
            DocumentHolder docToInsert1 = new DocumentHolder(dataDir + "AppendixA.docx");
            DocumentHolder docToInsert2 = new DocumentHolder(dataDir + "AppendixB.docx");

            // Prepare the array of data sources and their corresponding names used in the template.
            object[] dataSources = { docToInsert1, docToInsert2 };
            string[] dataSourceNames = { "srcA", "srcB" };

            // Build the report again, this time providing the document sources.
            // The template should contain tags like <<doc [srcA.Document]>> and <<doc [srcB.Document]>>.
            engine.BuildReport(template, dataSources, dataSourceNames);

            // Save the final document. The format is inferred from the extension (DOCX).
            template.Save(dataDir + "FinalReport.docx");
        }
    }
}
