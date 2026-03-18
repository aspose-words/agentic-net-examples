using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlMetadataExample
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new blank document and add a content control (SDT) that
            //    will hold custom metadata as plain‑text JSON.
            // -----------------------------------------------------------------
            Document doc = new Document();                         // create
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a block‑level plain‑text content control.
            StructuredDocumentTag metadataTag = new StructuredDocumentTag(
                doc, SdtType.PlainText, MarkupLevel.Block);

            // Use the Tag property as an identifier and the Title as a friendly name.
            metadataTag.Tag = "CustomMetadata";
            metadataTag.Title = "ArticleMetadata";

            // Prevent users from editing the metadata inside the control.
            metadataTag.LockContents = true;

            // Example metadata in JSON format.
            string jsonMetadata = "{\"id\":123,\"category\":\"news\",\"author\":\"John Doe\"}";

            // Insert the metadata text into the content control.
            metadataTag.AppendChild(new Run(doc, jsonMetadata));

            // Insert the content control into the document body.
            builder.InsertNode(metadataTag);

            // Save the document to disk.
            doc.Save("ArticleWithMetadata.docx");                  // save

            // -----------------------------------------------------------------
            // 2. Load the saved document and extract the metadata from the
            //    content control for further processing (e.g., indexing).
            // -----------------------------------------------------------------
            Document loadedDoc = new Document("ArticleWithMetadata.docx"); // load

            // Retrieve all StructuredDocumentTag nodes in the document.
            NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            // Collect extracted metadata strings.
            List<string> extractedMetadata = new List<string>();

            foreach (StructuredDocumentTag sdt in sdtNodes)
            {
                // Identify the correct content control by its Tag value.
                if (sdt.Tag == "CustomMetadata")
                {
                    // The inner text of the SDT contains the JSON metadata.
                    string metadata = sdt.GetText(); // includes the text inside the control
                    extractedMetadata.Add(metadata);
                }
            }

            // Output the extracted metadata (simulating indexing).
            foreach (string meta in extractedMetadata)
            {
                Console.WriteLine("Extracted metadata: " + meta);
                // Here you could parse the JSON and send fields to a search engine index.
            }
        }
    }
}
