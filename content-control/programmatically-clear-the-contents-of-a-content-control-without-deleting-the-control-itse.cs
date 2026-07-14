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

        // Insert an inline plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        sdt.Title = "SampleControl";
        sdt.Tag = "sample-tag";

        // Add some initial text inside the control.
        sdt.AppendChild(new Run(doc, "Initial content"));
        // Append the control to the first paragraph.
        builder.CurrentParagraph.AppendChild(sdt);

        // Save the document before clearing (optional, just for demonstration).
        doc.Save("BeforeClear.docx");

        // Locate the content control by its title.
        StructuredDocumentTag found = (StructuredDocumentTag)doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .FirstOrDefault(node => ((StructuredDocumentTag)node).Title == "SampleControl");

        if (found != null)
        {
            // Clear the contents of the control while keeping the control itself.
            found.Clear();
        }

        // Save the document after clearing.
        doc.Save("AfterClear.docx");
    }
}
