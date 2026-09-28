using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag linkControl = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        linkControl.Title = "LinkControl";
        linkControl.Tag = "link-control";

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(linkControl);

        // Inside the content control, create a paragraph that will hold the hyperlink.
        Paragraph innerParagraph = new Paragraph(doc);
        linkControl.AppendChild(innerParagraph);

        // Insert a hyperlink into the inner paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(innerParagraph);
        builder.InsertHyperlink("Click Here", "https://example.com", false);

        // Save the document as DOCX.
        const string docxPath = "output.docx";
        doc.Save(docxPath);

        // Convert the document to HTML.
        const string htmlPath = "output.html";
        doc.Save(htmlPath, SaveFormat.Html);

        // Load the generated HTML and verify that the hyperlink target URL is present.
        string htmlContent = File.ReadAllText(htmlPath);
        const string expectedUrl = "https://example.com";

        if (!htmlContent.Contains($"href=\"{expectedUrl}\""))
        {
            throw new InvalidOperationException($"The expected hyperlink URL '{expectedUrl}' was not found in the HTML output.");
        }

        // Indicate success.
        Console.WriteLine("Hyperlink embedded in content control and verified after conversion.");
    }
}
