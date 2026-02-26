using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Simple data class that holds a Document instance.
    public class DocumentTestClass
    {
        public Document Document { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template that contains the ReportingEngine tag <<doc [src.Document]>>.
            Document template = new Document("Template.docx");

            // Load the document that we want to insert into the template.
            Document source = new Document("Source.docx");

            // Prepare a collection of data sources using LINQ.
            // In this example we have only one document, but the same approach works for many.
            List<DocumentTestClass> dataSources = new List<DocumentTestClass>
            {
                new DocumentTestClass { Document = source }
            };

            // Create the reporting engine. Options can be adjusted as needed.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report by passing the template, the array of data sources,
            // and the corresponding names used in the template (here "src").
            engine.BuildReport(template, dataSources.ToArray(), new[] { "src" });

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
