using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlLockExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add some introductory text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("The following content control is locked (read‑only): ");

            // Create an inline plain‑text content control.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "LockedControl",
                Tag = "locked",
                // Prevent the user from editing the contents.
                LockContents = true,
                // Prevent the user from deleting the content control.
                LockContentControl = true
            };

            // Set the displayed text inside the content control.
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, "Read‑only content"));

            // Insert the content control into the document.
            builder.InsertNode(sdt);

            // Save the resulting document.
            doc.Save("LockedContentControl.docx");
        }
    }
}
