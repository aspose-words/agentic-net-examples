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

        // Add an introductory paragraph.
        builder.Writeln("Document with a rich‑text content control that will be filled with HTML:");

        // Create a block‑level rich‑text StructuredDocumentTag (content control).
        StructuredDocumentTag richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "HtmlContent",
            Tag = "html-content"
        };

        // Add a placeholder paragraph inside the content control.
        Paragraph placeholder = new Paragraph(doc);
        placeholder.AppendChild(new Run(doc, "Placeholder text"));
        richSdt.AppendChild(placeholder);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(richSdt);

        // Save the seed document (optional, shows the initial state).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document back (simulating a real file scenario).
        Document loadedDoc = new Document(inputPath);

        // Locate the rich‑text content control by its Title.
        StructuredDocumentTag targetSdt = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(s => s.Title == "HtmlContent");

        if (targetSdt == null)
        {
            Console.WriteLine("Content control not found.");
            return;
        }

        // Retrieve formatted HTML from a web service.
        string html = await GetHtmlFromWebAsync();

        // Replace the existing contents of the content control with the HTML.
        targetSdt.RemoveAllChildren();                     // Clear any existing child nodes.
        Paragraph insertionParagraph = new Paragraph(loadedDoc);
        targetSdt.AppendChild(insertionParagraph);         // The HTML will be inserted into this paragraph.

        DocumentBuilder sdtBuilder = new DocumentBuilder(loadedDoc);
        sdtBuilder.MoveTo(insertionParagraph);
        sdtBuilder.InsertHtml(html);                       // Parses the HTML and adds formatted content.

        // Save the updated document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        Console.WriteLine($"HTML inserted and document saved to '{outputPath}'.");
    }

    // Retrieves a simple HTML snippet from a public web service.
    private static async Task<string> GetHtmlFromWebAsync()
    {
        // Example endpoint that returns a small HTML page.
        const string url = "https://httpbin.org/html";

        using HttpClient client = new HttpClient();
        HttpResponseMessage response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();

        // The response body contains HTML.
        string html = await response.Content.ReadAsStringAsync();

        // For demonstration, we could trim the outer <html> tags if desired,
        // but InsertHtml can handle a full HTML document.
        return html;
    }
}
