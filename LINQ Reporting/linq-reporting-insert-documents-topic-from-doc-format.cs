using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the template that contains <<doc [src.Document]>> tags.
        Document template = new Document("Template.docx");

        // Get all DOC files that should be inserted into the template.
        string[] docFiles = Directory.GetFiles("DocsToInsert", "*.doc");

        // Create a collection of wrapper objects that expose a Document property.
        // This mirrors the example in Aspose.Words documentation.
        var sources = docFiles.Select(f => new DocumentWrapper
        {
            Document = new Document(f) // Load each source document.
        }).ToArray();

        // Prepare the data sources and their names for the ReportingEngine.
        object[] dataSources = sources.Cast<object>().ToArray();
        string[] dataSourceNames = sources.Select((s, i) => $"src{i}").ToArray();

        // Build the report – the engine will replace the <<doc>> tags with the loaded documents.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSources, dataSourceNames);

        // Save the populated document.
        template.Save("Result.docx");
    }

    // Simple wrapper class required by the ReportingEngine to expose the document.
    public class DocumentWrapper
    {
        public Document Document { get; set; }
    }
}
