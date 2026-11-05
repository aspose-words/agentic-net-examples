using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Loading;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample document with a block‑level rich‑text content control.
        Document doc = new Document();
        StructuredDocumentTag richSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "HtmlContent",
            Tag = "html-content"
        };

        // Add a placeholder paragraph inside the content control.
        Paragraph placeholderPara = new Paragraph(doc);
        placeholderPara.AppendChild(new Run(doc, "Placeholder text."));
        richSdt.AppendChild(placeholderPara);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(richSdt);

        // Optional: save the seed document for inspection.
        doc.Save("input.docx");

        // 2. Retrieve formatted HTML from a web service.
        string html;
        using (HttpClient client = new HttpClient())
        {
            HttpResponseMessage response = client.GetAsync("https://www.example.com").Result;
            response.EnsureSuccessStatusCode();
            html = response.Content.ReadAsStringAsync().Result;
        }

        // 3. Locate the rich‑text content control by its title.
        StructuredDocumentTag targetSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(s => s.Title == "HtmlContent");

        if (targetSdt == null)
            throw new InvalidOperationException("Rich‑text content control not found.");

        // 4. Remove any existing children of the content control.
        targetSdt.RemoveAllChildren();

        // 5. Load the retrieved HTML into a temporary document.
        using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
            Document htmlDoc = new Document(htmlStream, loadOptions);

            // 6. Import only block‑level nodes (paragraphs, tables, etc.) into the content control.
            var blockNodes = htmlDoc.FirstSection.Body.GetChildNodes(NodeType.Any, true)
                .Where(n => n.NodeType == NodeType.Paragraph || n.NodeType == NodeType.Table);

            foreach (Node node in blockNodes)
            {
                Node importedNode = doc.ImportNode(node, true, ImportFormatMode.KeepSourceFormatting);
                targetSdt.AppendChild(importedNode);
            }
        }

        // 7. Serialize some information about the updated control to JSON.
        var info = new
        {
            Title = targetSdt.Title,
            Tag = targetSdt.Tag,
            Content = targetSdt.GetText().Trim()
        };
        string json = JsonConvert.SerializeObject(info, Formatting.Indented);
        File.WriteAllText("control-info.json", json);

        // 8. Save the updated document.
        doc.Save("output.docx");
    }
}
