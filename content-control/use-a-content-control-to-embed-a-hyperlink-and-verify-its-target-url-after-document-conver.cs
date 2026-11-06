using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag linkControl = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "LinkControl",
            Tag = "link-control"
        };

        // Add a paragraph that will hold the hyperlink.
        Paragraph hyperlinkParagraph = new Paragraph(doc);
        linkControl.AppendChild(hyperlinkParagraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(linkControl);

        // Use DocumentBuilder to insert a hyperlink into the paragraph inside the content control.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(hyperlinkParagraph);
        builder.InsertHyperlink("https://example.com", "Click Here", false);

        // Ensure fields are up‑to‑date before saving.
        doc.UpdateFields();

        // Save the original DOCX.
        const string docxPath = "hyperlink-sdt.docx";
        doc.Save(docxPath);

        // Convert the document to HTML.
        const string htmlPath = "hyperlink-sdt.html";
        doc.Save(htmlPath, SaveFormat.Html);

        // Verify that the hyperlink target URL is present in the HTML output.
        string htmlContent = File.ReadAllText(htmlPath);
        const string expectedUrl = "https://example.com";
        if (!htmlContent.Contains(expectedUrl))
        {
            throw new InvalidOperationException("Hyperlink target URL was not preserved after conversion.");
        }

        // Indicate success.
        Console.WriteLine("Hyperlink target URL verified successfully.");
    }
}
