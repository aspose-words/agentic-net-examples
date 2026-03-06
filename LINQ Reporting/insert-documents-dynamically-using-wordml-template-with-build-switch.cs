using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDynamicInsert
{
    // Simple data source class that holds a Document to be inserted.
    public class DocumentSource
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the WORDML (or DOCX) template that contains the build switch tag, e.g. <<doc [src.Document]>>.
            Document template = new Document("Template.docx");

            // Load the document that we want to insert dynamically.
            Document docToInsert = new Document("Source.docx");

            // Prepare the data source object.
            DocumentSource src = new DocumentSource { Document = docToInsert };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the template and the data source.
            // The name "src" matches the name used in the template tag.
            engine.BuildReport(template, new object[] { src }, new[] { "src" });

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
