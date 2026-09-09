using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

namespace ContentControlMetadataExample
{
    public class Program
    {
        public static void Main()
        {
            // Path for the generated files.
            const string docPath = "metadata.docx";
            const string jsonPath = "metadata.json";

            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Prepare a paragraph to host the content controls.
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

            // 3. Define metadata items to store.
            var metadataItems = new Dictionary<string, string>
            {
                { "ProductId", "12345" },
                { "Category", "Electronics" },
                { "Price", "199.99" }
            };

            // 4. Insert a plain‑text content control for each metadata item.
            foreach (var kvp in metadataItems)
            {
                // Create an inline plain‑text StructuredDocumentTag.
                StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
                {
                    Title = kvp.Key,          // Use the key as the title (friendly name).
                    Tag = kvp.Key.ToLower()   // Use a lowercase tag for easy lookup.
                };

                // Clear any default children and set the initial value.
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(doc, kvp.Value));

                // Append the content control to the paragraph.
                paragraph.AppendChild(sdt);

                // Add a space after each control for readability.
                paragraph.AppendChild(new Run(doc, " "));
            }

            // 5. Save the document containing the metadata.
            doc.Save(docPath);

            // -----------------------------------------------------------------
            // 6. Load the document back and extract the metadata from the controls.
            Document loadedDoc = new Document(docPath);

            // Collect metadata from all StructuredDocumentTag nodes that have a Title.
            var extractedMetadata = new Dictionary<string, string>();
            NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            foreach (StructuredDocumentTag sdt in sdtNodes)
            {
                if (!string.IsNullOrEmpty(sdt.Title))
                {
                    // Get the text inside the content control and trim whitespace.
                    string value = sdt.GetText().Trim();
                    extractedMetadata[sdt.Title] = value;
                }
            }

            // 7. Serialize the extracted metadata to JSON.
            string json = JsonConvert.SerializeObject(extractedMetadata, Formatting.Indented);

            // 8. Save the JSON to a file.
            File.WriteAllText(jsonPath, json);

            // Optional: write the JSON to console (no interactive input required).
            Console.WriteLine("Extracted metadata JSON:");
            Console.WriteLine(json);
        }
    }
}
