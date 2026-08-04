using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json; // Included as required by the task, though not used.

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to host the content control.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create an inline plain‑text content control.
        StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        contentControl.Title = "SampleControl";
        contentControl.Tag = "sample-tag";

        // Add some initial text inside the control.
        contentControl.RemoveAllChildren(); // Ensure it is empty before adding text.
        contentControl.AppendChild(new Run(doc, "Initial content inside the control."));

        // Insert the content control into the paragraph.
        paragraph.AppendChild(contentControl);

        // Save the document with the populated content control.
        string initialPath = Path.Combine(Environment.CurrentDirectory, "initial.docx");
        doc.Save(initialPath);

        // Locate the content control by its title.
        StructuredDocumentTag foundControl = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                                .OfType<StructuredDocumentTag>()
                                                .FirstOrDefault(sdt => sdt.Title == "SampleControl");

        if (foundControl != null)
        {
            // Clear the contents of the control while keeping the control itself.
            foundControl.Clear();
        }

        // Save the document after clearing the control's contents.
        string clearedPath = Path.Combine(Environment.CurrentDirectory, "cleared.docx");
        doc.Save(clearedPath);
    }
}
