using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some introductory text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Below is a locked content control (read‑only).");

        // Insert an inline plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "ReadOnlyControl",
            Tag = "readonly",
            // Prevent the user from editing the contents.
            LockContents = true,
            // Prevent the user from deleting the content control.
            LockContentControl = true
        };

        // Add placeholder text inside the content control.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "This text cannot be edited or removed."));

        // Insert the content control into the document.
        builder.InsertNode(sdt);

        // Save the resulting document.
        doc.Save("LockedContentControl.docx");
    }
}
