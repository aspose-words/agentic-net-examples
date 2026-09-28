using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a paragraph to host the inline content control.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        if (paragraph == null)
        {
            paragraph = new Paragraph(doc);
            doc.FirstSection.Body.AppendChild(paragraph);
        }

        // Create an inline plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        sdt.Title = "CustomerName";   // Set the title for identification.
        sdt.Tag = "customer-name";    // Set the tag for identification.

        // Add placeholder text inside the content control.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Contoso"));

        // Insert the content control into the paragraph.
        paragraph.AppendChild(sdt);

        // Save the resulting document.
        doc.Save("ContentControlTitleTag.docx");
    }
}
