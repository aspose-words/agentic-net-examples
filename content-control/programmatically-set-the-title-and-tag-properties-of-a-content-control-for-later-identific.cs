using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to host the inline content control.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph ?? new Paragraph(doc);
        if (doc.FirstSection.Body.FirstParagraph == null)
        {
            doc.FirstSection.Body.AppendChild(paragraph);
        }

        // Create an inline plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

        // Set the Title and Tag properties for later identification.
        sdt.Title = "CustomerName";
        sdt.Tag = "customer-name";

        // Add some placeholder text inside the content control.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Contoso"));

        // Insert the content control into the paragraph.
        paragraph.AppendChild(sdt);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ContentControlTitleTag.docx");
        doc.Save(outputPath);
    }
}
