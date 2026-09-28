using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        contentControl.Title = "ReadOnlySection";

        // Lock the content control so the user cannot edit its contents or its properties.
        contentControl.LockContents = true;
        contentControl.LockContentControl = true;

        // Add some sample text inside the locked content control.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "This content is read‑only."));
        contentControl.AppendChild(paragraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(contentControl);

        // Save the resulting document.
        doc.Save("LockedContentControl.docx");
    }
}
