using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph of the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create an inline plain‑text content control.
        StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        contentControl.Title = "ReadOnlyInfo";
        contentControl.Tag = "readonly-info";

        // Set the text inside the content control.
        contentControl.RemoveAllChildren();
        contentControl.AppendChild(new Run(doc, "This text is read‑only."));

        // Lock the content control to make it read‑only.
        contentControl.LockContents = true;          // Prevent editing of the content.
        contentControl.LockContentControl = true;    // Prevent deletion or moving of the control.

        // Insert the locked content control into the paragraph.
        paragraph.AppendChild(contentControl);

        // Save the resulting document.
        const string outputFileName = "locked-content-control.docx";
        doc.Save(outputFileName);
    }
}
