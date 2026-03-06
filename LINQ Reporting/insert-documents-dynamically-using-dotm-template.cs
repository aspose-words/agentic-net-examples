using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDemo
{
    // Simple class that will be used as a data source for the reporting engine.
    // The property name must match the field used in the DOTM template.
    public class DocumentDataSource
    {
        // The reporting engine will look for a property named "Document"
        // when it encounters the tag <<doc [src.Document]>> in the template.
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTM template that contains a reporting tag like:
            // <<doc [src.Document]>>
            // The template can also contain other merge fields.
            Document template = new Document("Template.dotm");

            // Load the document that we want to insert dynamically.
            Document documentToInsert = new Document("SubDocument.docx");

            // Prepare the data source instance.
            var dataSource = new DocumentDataSource
            {
                Document = documentToInsert
            };

            // Create the reporting engine and populate the template.
            // The third argument ("src") is the name used in the template tag.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "src");

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
