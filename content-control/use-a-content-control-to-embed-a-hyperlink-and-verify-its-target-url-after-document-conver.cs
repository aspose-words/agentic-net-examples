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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a block‑level rich‑text content control into the document body.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        doc.FirstSection.Body.AppendChild(sdt);

        // Add a paragraph inside the content control to host the hyperlink.
        Paragraph innerParagraph = new Paragraph(doc);
        sdt.AppendChild(innerParagraph);
        builder.MoveTo(innerParagraph);

        // Insert a hyperlink inside the content control.
        const string linkText = "Visit Example";
        const string url = "https://example.com";
        builder.InsertHyperlink(linkText, url, false);

        // Save the document as DOCX.
        const string docxPath = "output.docx";
        doc.Save(docxPath);

        // Convert the document to HTML.
        const string htmlPath = "output.html";
        doc.Save(htmlPath, SaveFormat.Html);

        // Verify that the hyperlink target URL appears in the generated HTML.
        string htmlContent = File.ReadAllText(htmlPath);
        bool containsUrl = htmlContent.Contains(url, StringComparison.OrdinalIgnoreCase);

        // Output verification result.
        Console.WriteLine(containsUrl
            ? "Hyperlink verified: target URL found in HTML output."
            : "Verification failed: target URL not found in HTML output.");
    }
}
