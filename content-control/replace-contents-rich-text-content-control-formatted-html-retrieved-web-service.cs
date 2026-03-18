using System;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static async Task Main()
    {
        // Load the source document that contains a rich‑text content control (SDT).
        Document doc = new Document("Input.docx");

        // Locate the first StructuredDocumentTag (rich‑text content control) in the document.
        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
        if (sdt == null)
            throw new InvalidOperationException("No StructuredDocumentTag found in the document.");

        // Retrieve formatted HTML from a web service.
        string html = await GetHtmlAsync("https://example.com/api/content");

        // Remove any existing content inside the content control.
        sdt.Clear();

        // Position the DocumentBuilder at the start of the cleared control.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Use MoveTo(sdt) – StructuredDocumentTag does not expose FirstParagraph.
        builder.MoveTo(sdt);

        // Insert the HTML. The InsertHtml method parses the HTML and applies its formatting.
        builder.InsertHtml(html);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Helper method to download HTML from a URL.
    private static async Task<string> GetHtmlAsync(string requestUri)
    {
        using HttpClient client = new HttpClient();
        return await client.GetStringAsync(requestUri);
    }
}
