using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (always exists in a new document).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create an inline plain‑text content control.
        StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        contentControl.Title = "SampleControl";
        contentControl.Tag = "sample-tag";

        // Add some initial text to the control.
        contentControl.AppendChild(new Run(doc, "Initial content"));
        paragraph.AppendChild(contentControl);

        // Save the document before clearing (optional, shows the original state).
        doc.Save("before-clear.docx");

        // Clear the contents of the content control while keeping the control itself.
        contentControl.Clear();

        // Save the document after clearing.
        doc.Save("after-clear.docx");
    }
}
