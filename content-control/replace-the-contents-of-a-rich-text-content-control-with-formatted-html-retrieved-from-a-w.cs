using System;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a block‑level rich‑text content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "HtmlContent",
            Tag = "html-content"
        };

        // Add a placeholder paragraph (will be removed later).
        Paragraph placeholder = new Paragraph(doc);
        placeholder.AppendChild(new Run(doc, "Placeholder text"));
        richTextSdt.AppendChild(placeholder);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Retrieve formatted HTML from a web service.
        string html = GetHtmlFromWeb().GetAwaiter().GetResult();

        // Clear existing placeholder and add an empty paragraph to host the HTML.
        richTextSdt.RemoveAllChildren();
        Paragraph hostParagraph = new Paragraph(doc);
        richTextSdt.AppendChild(hostParagraph);

        // Position the builder inside the newly added paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(hostParagraph);
        builder.InsertHtml(html); // Insert the HTML; formatting is preserved.

        // Save the resulting document.
        doc.Save("output.docx");
    }

    // Simple helper that downloads HTML from a public URL.
    private static async Task<string> GetHtmlFromWeb()
    {
        const string url = "https://www.example.com"; // Any page that returns HTML.
        using HttpClient client = new HttpClient();
        HttpResponseMessage response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }
}
