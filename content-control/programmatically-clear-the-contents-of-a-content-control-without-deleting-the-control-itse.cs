using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlClearExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a plain‑text content control (inline level) into the document.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "SampleControl",
                Tag = "sample-control"
            };
            builder.InsertNode(sdt);

            // Add some initial text inside the content control.
            sdt.AppendChild(new Run(doc, "Initial content inside the control."));

            // Clear the contents of the content control while keeping the control itself.
            sdt.Clear();

            // Save the resulting document.
            doc.Save("ClearedContentControl.docx");

            // Indicate completion (no interactive input required).
            Console.WriteLine("Document saved as ClearedContentControl.docx");
        }
    }
}
