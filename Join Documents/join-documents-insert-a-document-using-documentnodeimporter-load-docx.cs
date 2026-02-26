using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the destination document (the document that will receive the inserted content).
        Document dstDoc = new Document("Destination.docx");

        // Load the source document (the document whose content will be inserted).
        Document srcDoc = new Document("Source.docx");

        // Choose the node after which the source content will be inserted.
        // In this example we insert after the first paragraph of the destination document.
        Node insertionDestination = dstDoc.FirstSection.Body.FirstParagraph;

        // Ensure the destination node is a Paragraph or Table as required by the importer logic.
        if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
            throw new ArgumentException("The destination node should be either a paragraph or table.");

        // The parent node that will receive the imported nodes.
        CompositeNode destinationParent = insertionDestination.ParentNode;

        // Create a NodeImporter to efficiently import nodes from srcDoc to dstDoc.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Iterate through all block-level nodes in each section of the source document.
        foreach (Section srcSection in srcDoc.Sections)
        {
            foreach (Node srcNode in srcSection.Body)
            {
                // Skip the last empty paragraph of a section (Word adds this automatically).
                if (srcNode.NodeType == NodeType.Paragraph)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.IsEndOfSection && !para.HasChildNodes)
                        continue;
                }

                // Import the node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionDestination);
                insertionDestination = importedNode; // Update the insertion point for the next node.
            }
        }

        // Save the combined document.
        dstDoc.Save("Combined.docx");
    }
}
