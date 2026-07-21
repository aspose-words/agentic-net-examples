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

        // Use DocumentBuilder to insert a paragraph that contains a hyperlink.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertParagraph(); // Ensure we are in a new paragraph.
        builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
        builder.Writeln(); // End the paragraph with a line break.

        // Retrieve the paragraph that now contains the hyperlink.
        Paragraph linkParagraph = builder.CurrentParagraph;

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag sdt = new StructuredDocumentTag(
            doc,
            SdtType.RichText,
            MarkupLevel.Block)
        {
            Title = "LinkControl",
            Tag = "link-control"
        };

        // Move the hyperlink paragraph into the content control.
        sdt.AppendChild(linkParagraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Save the DOCX file.
        const string docxPath = "hyperlink_sdt.docx";
        doc.Save(docxPath);

        // Convert the document to HTML.
        const string htmlPath = "hyperlink_sdt.html";
        doc.Save(htmlPath, SaveFormat.Html);

        // Verify that the hyperlink URL is present in the HTML output.
        string htmlContent = File.ReadAllText(htmlPath);
        bool containsUrl = htmlContent.Contains("https://www.aspose.com");

        Console.WriteLine(containsUrl
            ? "Verification succeeded: hyperlink URL found in HTML."
            : "Verification failed: hyperlink URL not found in HTML.");
    }
}
