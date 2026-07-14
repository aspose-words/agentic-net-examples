using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create an inline rich‑text content control (SDT).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Inline)
        {
            Title = "HyperlinkControl",
            Tag = "hyperlink"
        };

        // Insert the content control into the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(sdt);

        // Insert a hyperlink directly inside the inline content control.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(sdt);
        builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);

        // Save the document as DOCX.
        const string docxPath = "HyperlinkControl.docx";
        doc.Save(docxPath);

        // Load the saved document.
        Document loadedDoc = new Document(docxPath);

        // Locate the first hyperlink field in the document.
        FieldHyperlink? hyperlinkField = loadedDoc.Range.Fields
            .OfType<FieldHyperlink>()
            .FirstOrDefault();

        // Verify that the hyperlink target URL matches the expected value.
        const string expectedUrl = "https://www.aspose.com";
        bool isUrlCorrect = hyperlinkField != null && hyperlinkField.Address == expectedUrl;

        // Output the verification result.
        Console.WriteLine(isUrlCorrect
            ? $"Hyperlink target verified: {hyperlinkField!.Address}"
            : "Hyperlink target verification failed.");
    }
}
