using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Loading;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document that contains a block‑level rich‑text
        //    content control. This document will act as the input file.
        // -----------------------------------------------------------------
        Document doc = new Document();
        StructuredDocumentTag richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richSdt.Title = "HtmlContent";
        richSdt.Tag = "html-content";

        // The content control must contain at least one paragraph.
        Paragraph placeholder = new Paragraph(doc);
        placeholder.AppendChild(new Run(doc, "Placeholder"));
        richSdt.AppendChild(placeholder);

        // Insert the content control into the document body and save.
        doc.FirstSection.Body.AppendChild(richSdt);
        doc.Save("input.docx");

        // -----------------------------------------------------------------
        // 2. Load the document that contains the content control.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document("input.docx");

        // Locate the rich‑text content control by its Title.
        StructuredDocumentTag? targetSdt = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(s => s.Title == "HtmlContent");

        if (targetSdt != null)
        {
            // -----------------------------------------------------------------
            // 3. Retrieve formatted HTML from a web service.
            // -----------------------------------------------------------------
            string html = GetHtmlFromWebAsync().GetAwaiter().GetResult();

            // -----------------------------------------------------------------
            // 4. Remove any existing children of the content control.
            // -----------------------------------------------------------------
            targetSdt.RemoveAllChildren();

            // -----------------------------------------------------------------
            // 5. Load the HTML into a temporary Aspose.Words document.
            // -----------------------------------------------------------------
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
                Document htmlDoc = new Document(htmlStream, loadOptions);

                // -----------------------------------------------------------------
                // 6. Import only block‑level nodes (Paragraphs, Tables, etc.) into the
                //    content control. Inline nodes such as Run cannot be appended
                //    directly to a StructuredDocumentTag.
                // -----------------------------------------------------------------
                foreach (Paragraph para in htmlDoc.FirstSection.Body.Paragraphs)
                {
                    Node imported = loadedDoc.ImportNode(para, true, ImportFormatMode.KeepSourceFormatting);
                    targetSdt.AppendChild(imported);
                }
            }
        }

        // -----------------------------------------------------------------
        // 7. Save the updated document.
        // -----------------------------------------------------------------
        loadedDoc.Save("output.docx");
    }

    private static async Task<string> GetHtmlFromWebAsync()
    {
        using (HttpClient client = new HttpClient())
        {
            // Example URL that returns HTML content.
            HttpResponseMessage response = await client.GetAsync("https://www.example.com");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
    }
}
