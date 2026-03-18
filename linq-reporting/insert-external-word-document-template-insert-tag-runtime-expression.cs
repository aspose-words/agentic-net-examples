using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InsertExternalDocumentExample
{
    // Simple data source class that will be referenced from the template.
    // The template should contain a tag like <<doc [src.Document]>> where "src" is the name we give to the data source.
    public class DocumentSource
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the template that contains the insert tag.
            // Example tag inside the template: <<doc [src.Document]>>
            Document template = new Document("Template.docx");

            // Load the external Word document that we want to insert.
            Document externalDoc = new Document("External.docx");

            // Prepare the data source instance.
            var src = new DocumentSource { Document = externalDoc };

            // Use the ReportingEngine to process the template.
            // The second parameter is the name that will be used inside the tag (src in this case).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, src, "src");

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
