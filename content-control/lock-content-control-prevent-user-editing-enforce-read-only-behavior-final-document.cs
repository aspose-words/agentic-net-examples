using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline plain‑text structured document tag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

        // Prohibit the user from editing the contents of the control.
        sdt.LockContents = true;

        // Prohibit the user from deleting the control itself.
        sdt.LockContentControl = true;

        // Optional: give the control a title and tag for identification.
        sdt.Title = "ReadOnlyControl";
        sdt.Tag = "ReadOnlyTag";

        // Add some explanatory text before the control.
        builder.Writeln("The following content control is read‑only and cannot be deleted:");

        // Insert the locked content control into the document.
        builder.InsertNode(sdt);

        // Save the resulting document.
        doc.Save("LockedContentControl.docx");
    }
}
