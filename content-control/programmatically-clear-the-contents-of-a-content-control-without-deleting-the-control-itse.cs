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

        // Write a line before the content control for clarity.
        builder.Writeln("Document with a plain‑text content control:");

        // Create an inline plain‑text StructuredDocumentTag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleControl",
            Tag = "sample"
        };

        // Add initial text inside the content control.
        sdt.AppendChild(new Run(doc, "Initial content"));

        // Append the content control to the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(sdt);

        // Save the document before clearing the control's contents.
        doc.Save("BeforeClear.docx");

        // Clear the contents of the content control while keeping the control itself.
        sdt.Clear();

        // Save the document after clearing the control's contents.
        doc.Save("AfterClear.docx");
    }
}
