using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph of the document body.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Create an inline plain‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

        // Set the Title and Tag properties for later identification.
        sdt.Title = "CustomerName";
        sdt.Tag = "customer-name";

        // Clear any default children and add a run with placeholder text.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Contoso"));

        // Insert the content control into the paragraph.
        firstParagraph.AppendChild(sdt);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ContentControlTitleTag.docx");
        doc.Save(outputPath);
    }
}
