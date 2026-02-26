using System;
using Aspose.Words;

namespace DocumentJoinExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source and destination DOCX files.
            string sourcePath = "Source.docx";
            string destinationPath = "Destination.docx";
            string resultPath = "JoinedDocument.docx";

            // Load the documents from disk.
            Document srcDoc = new Document(sourcePath);
            Document dstDoc = new Document(destinationPath);

            // Insert the source document using three different ImportFormatMode values.
            // 1. Use destination styles (source formatting may be altered).
            InsertDocumentUsingImporter(dstDoc, srcDoc, ImportFormatMode.UseDestinationStyles);

            // 2. Keep source formatting (styles are copied or converted to direct formatting).
            InsertDocumentUsingImporter(dstDoc, srcDoc, ImportFormatMode.KeepSourceFormatting);

            // 3. Keep different styles only (identical styles are reused, different ones are copied).
            InsertDocumentUsingImporter(dstDoc, srcDoc, ImportFormatMode.KeepDifferentStyles);

            // Save the combined document.
            dstDoc.Save(resultPath);
        }

        /// <summary>
        /// Inserts all block‑level nodes from srcDoc into dstDoc after the last node of the destination.
        /// The import respects the specified ImportFormatMode.
        /// </summary>
        static void InsertDocumentUsingImporter(Document dstDoc, Document srcDoc, ImportFormatMode mode)
        {
            // Create a NodeImporter that will handle style and list translation.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, mode);

            // Destination body where we will append the imported nodes.
            CompositeNode dstBody = dstDoc.LastSection.Body;

            // Iterate through each section of the source document.
            foreach (Section srcSection in srcDoc.Sections)
            {
                // Iterate through each node in the section's body (paragraphs, tables, etc.).
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the final empty paragraph that Word adds to each section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // Import the node into the destination document.
                    Node importedNode = importer.ImportNode(srcNode, true);

                    // Append the imported node to the destination body.
                    dstBody.AppendChild(importedNode);
                }
            }
        }
    }
}
