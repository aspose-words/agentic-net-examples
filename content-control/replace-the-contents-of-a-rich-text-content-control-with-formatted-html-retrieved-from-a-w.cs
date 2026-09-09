using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static async Task Main(string[] args)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Document before the content control.");

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "HtmlPlaceholder",
            Tag = "html-placeholder"
        };
        // The SDT must contain at least one paragraph to host content.
        Paragraph placeholderParagraph = new Paragraph(doc);
        richSdt.AppendChild(placeholderParagraph);
        doc.FirstSection.Body.AppendChild(richSdt);

        // Add some text after the control.
        builder.Writeln("Document after the content control.");

        // Optional: save the seed document (useful for debugging).
        doc.Save("seed.docx");

        // Retrieve formatted HTML from a web service.
        string htmlContent;
        using (HttpClient httpClient = new HttpClient())
        {
            // Example URL that returns a simple HTML page.
            HttpResponseMessage response = await httpClient.GetAsync("https://httpbin.org/html");
            response.EnsureSuccessStatusCode();
            htmlContent = await response.Content.ReadAsStringAsync();
        }

        // Locate the rich‑text content control by its title.
        StructuredDocumentTag targetSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(s => s.Title == "HtmlPlaceholder");

        if (targetSdt == null)
        {
            throw new InvalidOperationException("The target content control was not found.");
        }

        // Ensure we are working with a rich‑text control.
        if (targetSdt.SdtType != SdtType.RichText)
        {
            throw new InvalidOperationException("The target content control is not a rich‑text control.");
        }

        // Remove any existing children (placeholder text, etc.).
        targetSdt.RemoveAllChildren();

        // Insert a new paragraph that will receive the HTML.
        Paragraph htmlParagraph = new Paragraph(doc);
        targetSdt.AppendChild(htmlParagraph);

        // Move the builder to the new paragraph inside the SDT and insert the HTML.
        DocumentBuilder htmlBuilder = new DocumentBuilder(doc);
        htmlBuilder.MoveTo(htmlParagraph);
        htmlBuilder.InsertHtml(htmlContent);

        // Save the resulting document.
        doc.Save("output.docx");
    }
}
