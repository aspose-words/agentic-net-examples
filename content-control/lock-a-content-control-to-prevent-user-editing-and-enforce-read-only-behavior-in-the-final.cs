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

        // Add a paragraph describing the content control.
        builder.Writeln("Document with a locked content control:");

        // Create an inline plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "ReadOnlyControl",
            Tag = "readonly",
            // Prevent the user from editing the contents.
            LockContents = true,
            // Prevent the user from deleting the content control.
            LockContentControl = true
        };

        // Set the displayed text inside the control.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "This text cannot be edited or the control deleted."));

        // Insert the locked content control into the document.
        builder.InsertNode(sdt);

        // Save the resulting document.
        doc.Save("LockedContentControl.docx");
    }
}
