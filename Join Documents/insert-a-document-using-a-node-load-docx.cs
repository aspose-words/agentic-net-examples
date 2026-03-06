using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the destination document (the one we will insert into).
        Document dstDoc = new Document("Destination.docx");

        // Load the source document (the one we want to insert).
        Document srcDoc = new Document("Source.docx");

        // Choose the node after which the source content will be inserted.
        // For this example we use the first paragraph of the destination document.
        Node insertionNode = dstDoc.FirstSection.Body.FirstParagraph;

        // The insertion node must be a paragraph or a table.
        if (insertionNode.NodeType != NodeType.Paragraph && insertionNode.NodeType != NodeType.Table)
            throw new ArgumentException("Insertion node must be a paragraph or a table.");

        // Create a NodeImporter that will handle style, list and other translation
        // from the source document to the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // The parent of the insertion node (a CompositeNode) will receive the imported nodes.
        CompositeNode destinationParent = insertionNode.ParentNode;

        // Loop through all block-level nodes in each section of the source document.
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

                // Import the node into the destination document.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Insert the imported node after the current insertion point.
                destinationParent.InsertAfter(importedNode, insertionNode);

                // Update the insertion point so subsequent nodes are inserted sequentially.
                insertionNode = importedNode;
            }
        }

        // Save the resulting document using the provided Save method.
        dstDoc.Save("Result.docx");
    }
}
