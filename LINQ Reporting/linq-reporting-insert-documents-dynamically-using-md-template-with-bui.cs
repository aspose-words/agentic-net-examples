using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple wrapper class that holds a Document instance.
    public class DocumentTestClass
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a template document that contains LINQ Reporting tags.
            // ------------------------------------------------------------
            Document template = new Document();                     // create blank document
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a tag that will be replaced with the content of src.Document.
            builder.Writeln("First insertion:");
            builder.Writeln("<<doc [src.Document]>>");

            // Insert a tag that will be replaced with the content of src.Document,
            // but keep its own numbering (sourceNumbering switch).
            builder.Writeln("Second insertion with separate numbering:");
            builder.Writeln("<<doc [src.Document] -sourceNumbering>>");

            // ------------------------------------------------------------
            // 2. Prepare the data source – a collection of objects each exposing a Document.
            // ------------------------------------------------------------
            var data = new List<DocumentTestClass>();

            for (int i = 1; i <= 2; i++)
            {
                // Create a simple document that will be inserted into the template.
                Document subDoc = new Document();
                DocumentBuilder subBuilder = new DocumentBuilder(subDoc);
                subBuilder.Writeln($"Document {i} content.");

                // Wrap it in the test class.
                data.Add(new DocumentTestClass { Document = subDoc });
            }

            // ------------------------------------------------------------
            // 3. Build the report using ReportingEngine.
            // ------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after tag removal.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // BuildReport overload that accepts multiple data sources.
            // We pass a single data source (the array of DocumentTestClass objects)
            // and give it the name "src" so the template can reference src.Document.
            engine.BuildReport(
                template,
                new object[] { data.ToArray() },   // dataSources
                new string[] { "src" }             // dataSourceNames
            );

            // ------------------------------------------------------------
            // 4. Save the generated document.
            // ------------------------------------------------------------
            template.Save("Result.docx");
        }
    }
}
