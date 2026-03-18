using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a plain‑text structured document tag (content control) at inline level.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

        // Set a friendly title – visible when the user hovers over the control in Word.
        sdt.Title = "CustomerName";

        // Set a tag – an arbitrary string used for programmatic identification (e.g., when exporting to PDF).
        sdt.Tag = "CustomerNameTag";

        // Optional: set a visual cue such as background color.
        sdt.Color = Color.LightBlue;

        // Insert the content control into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertNode(sdt);

        // Save the document.
        doc.Save("ContentControlWithTitleAndTag.docx");
    }
}
