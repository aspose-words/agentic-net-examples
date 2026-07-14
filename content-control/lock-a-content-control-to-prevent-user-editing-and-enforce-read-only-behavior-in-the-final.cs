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

        // Create a plain‑text inline content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        sdt.Title = "ReadOnlyControl";
        sdt.Tag = "readonly";

        // Lock the contents so the user cannot edit them.
        sdt.LockContents = true;
        // Lock the control itself so the user cannot delete it.
        sdt.LockContentControl = true;

        // Add text inside the locked content control.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "This text cannot be edited or removed."));

        // Insert the locked content control into the document.
        builder.Write("Locked content control: ");
        builder.InsertNode(sdt);
        builder.Writeln();

        // Save the resulting document.
        doc.Save("LockedContentControl.docx");
    }
}
