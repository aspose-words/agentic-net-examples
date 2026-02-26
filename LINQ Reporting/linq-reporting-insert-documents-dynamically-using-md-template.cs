using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple wrapper class that holds a Document.
    // The template will reference the document via the property "Document".
    public class DocumentWrapper
    {
        public Document Document { get; }

        public DocumentWrapper(Document doc)
        {
            Document = doc ?? throw new ArgumentNullException(nameof(doc));
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the LINQ Reporting template that contains <<doc [src.Document]>> tags.
            // The template can be created beforehand with the required tags.
            Document template = new Document("TemplateWithDocTags.docx");

            // Load the source document that we want to insert dynamically.
            Document sourceDoc = new Document("SourceDocument.docx");

            // Wrap the source document in a class that the ReportingEngine can use as a data source.
            DocumentWrapper srcWrapper = new DocumentWrapper(sourceDoc);

            // Create the ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the overload that accepts multiple data sources.
            // The first (and only) data source is named "src" – this name must match the tag in the template.
            engine.BuildReport(
                template,
                new object[] { srcWrapper },          // data sources array
                new[] { "src" }                       // corresponding names array
            );

            // Save the resulting document.
            template.Save("ResultDocument.docx");
        }
    }
}
