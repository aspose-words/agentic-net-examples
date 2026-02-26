using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class DocumentWrapper
{
    // Property name must match the tag used in the template (src.Document)
    public Document Document { get; set; }

    public DocumentWrapper(Document doc)
    {
        Document = doc;
    }
}

class Program
{
    static void Main()
    {
        // Paths to the files – adjust as needed for your environment.
        string sourcePath = "Source.docx";
        string outputPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple template document that contains a reporting tag.
        //    The tag <<doc [src.Document]>> tells the ReportingEngine to
        //    insert the document referenced by src.Document at this location.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("=== Begin of Report ===");
        builder.Writeln("<<doc [src.Document]>>"); // Insertion point
        builder.Writeln("=== End of Report ===");

        // -----------------------------------------------------------------
        // 2. Load the document that will be inserted.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Wrap the source document in a class so the engine can access it
        //    via the tag. The property name (Document) matches the field used
        //    in the tag (src.Document).
        // -----------------------------------------------------------------
        DocumentWrapper wrapper = new DocumentWrapper(sourceDoc);

        // -----------------------------------------------------------------
        // 4. Build the report. The engine replaces the tag with the contents
        //    of sourceDoc, preserving its formatting.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, new object[] { wrapper }, new[] { "src" });

        // -----------------------------------------------------------------
        // 5. Save the merged document in DOC format.
        // -----------------------------------------------------------------
        template.Save(outputPath, SaveFormat.Doc);
    }
}
