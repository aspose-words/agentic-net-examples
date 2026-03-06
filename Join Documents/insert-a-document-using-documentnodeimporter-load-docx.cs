using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination document (the document that will receive the content).
        Document dstDoc = new Document("Destination.docx"); // replace with actual path

        // Load the source document (the document whose content will be inserted).
        Document srcDoc = new Document("Source.docx"); // replace with actual path

        // Choose an insertion point – here we use the first paragraph of the destination.
        Node insertionDestination = dstDoc.FirstSection.Body.FirstParagraph;

        // Verify that the insertion point is a paragraph or a table.
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("Insertion point must be a paragraph or a table.");

        // The parent node where new nodes will be added.
        CompositeNode destinationParent = insertionDestination.ParentNode;

        // Create a NodeImporter for efficient repeated imports.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Loop through all block‑level nodes in the source document and import them.
        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the final empty paragraph of a section (Word adds it automatically).
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node (deep clone belonging to the destination document).
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode; // move the pointer forward.
            }
        }

        // Save the resulting document.
        dstDoc.Save("Combined.docx"); // replace with desired output path
    }
}
