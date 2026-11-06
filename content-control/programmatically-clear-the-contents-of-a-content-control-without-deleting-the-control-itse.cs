using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create an inline plain‑text content control with a title and tag.
        StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleControl",
            Tag = "sample"
        };
        // Add some initial text inside the control.
        contentControl.AppendChild(new Run(doc, "Initial content"));

        // Insert the content control into the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(contentControl);

        // Save the document that contains the populated content control.
        doc.Save("original.docx");

        // Locate the content control by its title.
        StructuredDocumentTag? targetControl = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(sdt => sdt.Title == "SampleControl");

        if (targetControl != null)
        {
            // Clear all child nodes of the content control without removing the control itself.
            targetControl.RemoveAllChildren();

            // Optionally, you could add placeholder text here if desired.
            // targetControl.AppendChild(new Run(doc, "Placeholder"));
        }

        // Save the document after clearing the content control's contents.
        doc.Save("cleared.docx");
    }
}
