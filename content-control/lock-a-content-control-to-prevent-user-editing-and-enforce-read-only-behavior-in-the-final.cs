using System;
using System.IO;
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
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a heading.
            builder.Writeln("Document with a read‑only content control:");

            // Create an inline plain‑text content control.
            StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
            // Prevent editing of the control's contents.
            contentControl.LockContents = true;
            // Prevent deletion of the control itself.
            contentControl.LockContentControl = true;

            // Add placeholder text inside the control.
            Run innerRun = new Run(doc, "Locked content");
            contentControl.AppendChild(innerRun);

            // Insert the content control into the document at the current position (inside a paragraph).
            builder.InsertNode(contentControl);
            builder.Writeln(); // Move to a new line after the control.

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "LockedContentControl.docx");
            doc.Save(outputPath);
        }
    }
}
