using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.BuildingBlocks;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a block‑level rich‑text content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "HtmlContent",
            Tag = "html-content"
        };

        // Add a placeholder paragraph inside the content control.
        Paragraph placeholder = new Paragraph(doc);
        placeholder.AppendChild(new Run(doc, "Placeholder"));
        richTextSdt.AppendChild(placeholder);

        // Append the content control to the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Retrieve formatted HTML from a simple web service.
        string html;
        using (HttpClient httpClient = new HttpClient())
        {
            // httpbin.org returns a small HTML page suitable for the example.
            html = httpClient.GetStringAsync("https://httpbin.org/html").Result;
        }

        // Replace the existing contents of the rich‑text content control with the retrieved HTML.
        richTextSdt.RemoveAllChildren();

        // Create a new paragraph that will host the inserted HTML.
        Paragraph htmlParagraph = new Paragraph(doc);
        richTextSdt.AppendChild(htmlParagraph);

        // Move the builder to the new paragraph and insert the HTML.
        builder.MoveTo(htmlParagraph);
        builder.InsertHtml(html);

        // (Optional) Serialize the fetched HTML to a JSON file using Newtonsoft.Json.
        string json = JsonConvert.SerializeObject(new { html }, Formatting.Indented);
        File.WriteAllText("html.json", json);

        // Save the resulting document.
        doc.Save("output.docx");
    }
}
