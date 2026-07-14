using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a block‑level rich‑text content control that will hold the HTML.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "HtmlContent",
            Tag = "html-content"
        };
        // A block‑level SDT must contain at least one paragraph.
        Paragraph placeholderParagraph = new Paragraph(doc);
        placeholderParagraph.AppendChild(new Run(doc, "Placeholder text"));
        richTextSdt.AppendChild(placeholderParagraph);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Retrieve formatted HTML from a web service.
        string html;
        using (HttpClient httpClient = new HttpClient())
        {
            // Example public URL that returns simple HTML.
            HttpResponseMessage response = httpClient.GetAsync("https://www.example.com").Result;
            response.EnsureSuccessStatusCode();
            html = response.Content.ReadAsStringAsync().Result;
        }

        // Find the previously created rich‑text content control by its title.
        StructuredDocumentTag targetSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(s => s.Title == "HtmlContent");

        if (targetSdt == null)
        {
            throw new InvalidOperationException("Rich‑text content control not found.");
        }

        // Clear existing contents of the content control.
        targetSdt.RemoveAllChildren();

        // Insert a new paragraph into the SDT to host the HTML.
        Paragraph htmlParagraph = new Paragraph(doc);
        targetSdt.AppendChild(htmlParagraph);

        // Use DocumentBuilder to insert the HTML inside the paragraph.
        DocumentBuilder htmlBuilder = new DocumentBuilder(doc);
        htmlBuilder.MoveTo(htmlParagraph);
        htmlBuilder.InsertHtml(html);

        // Save the resulting document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {Path.GetFullPath(outputPath)}");
    }
}
