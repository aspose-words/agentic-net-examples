using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOT (template) document that contains <<doc [src.Document]>> tags.
        Document template = new Document("TemplateWithDocTag.docx");

        // Load the source documents that will be inserted dynamically.
        Document docToInsert1 = new Document("DocToInsert1.docx");
        Document docToInsert2 = new Document("DocToInsert2.docx");

        // Wrap each source document in a simple class exposing a Document property.
        // The template will reference the property using the build switch, e.g. <<doc [src1.Document]>>.
        var src1 = new DocumentWrapper { Document = docToInsert1 };
        var src2 = new DocumentWrapper { Document = docToInsert2 };

        // Prepare the data source arrays for the ReportingEngine.
        object[] dataSources = new object[] { src1, src2 };
        string[] dataSourceNames = new string[] { "src1", "src2" };

        // Build the report – the engine will replace the <<doc>> tags with the actual documents.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSources, dataSourceNames);

        // Save the populated document.
        template.Save("Result.docx");
    }

    // Helper class used by the template to expose the document to be inserted.
    public class DocumentWrapper
    {
        public Document Document { get; set; }
    }
}
