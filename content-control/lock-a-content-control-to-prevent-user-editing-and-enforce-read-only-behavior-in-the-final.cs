using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write introductory text.
        builder.Writeln("Document with a locked content control:");

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

        // Add default text inside the control.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "This text cannot be edited or removed."));

        // Insert the content control into the document.
        builder.InsertNode(sdt);
        builder.Writeln(); // Move to a new line after the control.

        // Save the resulting document.
        doc.Save("LockedContentControl.docx");
    }
}
