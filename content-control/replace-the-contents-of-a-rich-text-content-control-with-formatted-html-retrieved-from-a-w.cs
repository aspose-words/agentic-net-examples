using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static async Task Main()
    {
        // Create a sample document with a block‑level rich‑text content control.
        Document doc = new Document();
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richTextSdt.Title = "HtmlContent";
        richTextSdt.Tag = "html-content";

        // Add an initial placeholder paragraph inside the content control.
        Paragraph placeholder = new Paragraph(doc);
        placeholder.AppendChild(new Run(doc, "Placeholder text"));
        richTextSdt.AppendChild(placeholder);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);
        doc.Save("template.docx");

        // Retrieve formatted HTML from a web service.
        string html = await GetHtmlAsync("https://www.example.com");

        // Find the rich‑text content control by its title.
        StructuredDocumentTag targetSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(s => s.Title == "HtmlContent");

        if (targetSdt == null)
        {
            throw new InvalidOperationException("Rich‑text content control not found.");
        }

        // Replace the existing content with the retrieved HTML.
        targetSdt.RemoveAllChildren();
        Paragraph htmlParagraph = new Paragraph(doc);
        targetSdt.AppendChild(htmlParagraph);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(htmlParagraph);
        builder.InsertHtml(html);

        // Optionally serialize the inserted HTML to a JSON file (demonstrates Newtonsoft.Json usage).
        var payload = new { SourceUrl = "https://www.example.com", HtmlContent = html };
        string json = JsonConvert.SerializeObject(payload, Formatting.Indented);
        File.WriteAllText("htmlPayload.json", json);

        // Save the updated document.
        doc.Save("output.docx");
    }

    private static async Task<string> GetHtmlAsync(string url)
    {
        using HttpClient client = new HttpClient();
        HttpResponseMessage response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }
}
