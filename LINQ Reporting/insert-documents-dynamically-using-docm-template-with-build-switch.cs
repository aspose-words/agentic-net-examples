using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDynamicInsert
{
    // Simple data source class that will be referenced from the DOCM template.
    public class DocData
    {
        // The document to be inserted. In the template use a tag like <<doc [src.Document]>> or
        // <<doc [src.Document] -sourceNumbering>> to control numbering behavior.
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains the build switch tags.
            Document template = new Document("Template.docm");

            // Load the document that should be inserted dynamically.
            Document documentToInsert = new Document("Insert.docx");

            // Prepare the data source object.
            DocData data = new DocData
            {
                Document = documentToInsert
            };

            // Create the reporting engine and populate the template.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter is the name used in the template to reference the data source.
            engine.BuildReport(template, data, "src");

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
