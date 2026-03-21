using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("This is a document with a content control:");

        // Insert a plain‑text content control (structured document tag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleContentControl"
            // PlaceholderName is omitted because it requires an existing building block.
        };
        builder.InsertNode(sdt);

        // Add text inside the content control.
        sdt.AppendChild(new Run(doc, "Content inside control"));

        // Configure XPS save options.
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            UpdateFields = false
        };

        // Save the document as XPS.
        doc.Save("OutputWithContentControls.xps", xpsOptions);
    }
}
