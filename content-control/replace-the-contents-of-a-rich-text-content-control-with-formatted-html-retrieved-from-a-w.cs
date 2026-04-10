using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlHtmlReplace
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a rich‑text content control (block level) and give it a title for later lookup.
            StructuredDocumentTag richTextTag = builder.InsertStructuredDocumentTag(SdtType.RichText);
            richTextTag.Title = "HtmlContent";
            // Optional placeholder text.
            richTextTag.IsShowingPlaceholderText = true;

            // Save the intermediate document (optional, demonstrates the source file).
            const string sourcePath = "SourceDocument.docx";
            doc.Save(sourcePath, SaveFormat.Docx);

            // Retrieve formatted HTML from a web service.
            string html;
            using (HttpClient httpClient = new HttpClient())
            {
                // Example URL – you can replace this with any endpoint that returns HTML.
                HttpResponseMessage response = httpClient.GetAsync("https://www.example.com").Result;
                response.EnsureSuccessStatusCode();
                html = response.Content.ReadAsStringAsync().Result;
            }

            // Load the document that contains the content control.
            Document loadedDoc = new Document(sourcePath);
            DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

            // Find the rich‑text content control by its title.
            StructuredDocumentTag targetTag = null;
            NodeCollection tags = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            foreach (Node node in tags)
            {
                StructuredDocumentTag sdt = (StructuredDocumentTag)node;
                if (sdt.Title == "HtmlContent")
                {
                    targetTag = sdt;
                    break;
                }
            }

            if (targetTag == null)
                throw new InvalidOperationException("The target content control was not found.");

            // Clear any existing contents of the content control.
            targetTag.Clear();

            // Move the builder cursor inside the content control.
            loadedBuilder.MoveTo(targetTag);

            // Insert the retrieved HTML. Use builder formatting as base formatting.
            loadedBuilder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting);

            // Save the resulting document.
            const string resultPath = "ResultDocument.docx";
            loadedDoc.Save(resultPath, SaveFormat.Docx);
        }
    }
}
