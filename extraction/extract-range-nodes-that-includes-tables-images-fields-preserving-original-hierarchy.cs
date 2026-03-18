using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class ExtractSpecificNodes
{
    static void Main()
    {
        // Load the source document.
        Document srcDoc = new Document("input.docx");

        // Create an empty destination document.
        Document dstDoc = new Document();

        // Ensure the destination has at least one section and body.
        // (A new Document already contains a section with a body.)
        CompositeNode dstBody = dstDoc.FirstSection.Body;

        // Prepare an importer to copy nodes from source to destination.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Keep track of nodes that have already been scheduled for import to avoid duplicates.
        HashSet<Node> imported = new HashSet<Node>();

        // List to hold nodes in the order they appear in the source document.
        List<Node> nodesToImport = new List<Node>();

        // Walk through all nodes in the main story (body) in document order.
        foreach (Node node in srcDoc.FirstSection.Body.GetChildNodes(NodeType.Any, true))
        {
            Node topNode = null;

            // Determine the top‑level node that should be imported for each type.
            switch (node.NodeType)
            {
                case NodeType.Table:
                    topNode = node; // Table is a block node; import it directly.
                    break;

                case NodeType.Shape:
                    // Images are Shape nodes. Import the containing paragraph to keep hierarchy.
                    topNode = node.GetAncestor(NodeType.Paragraph);
                    break;

                case NodeType.FieldStart:
                    // Fields are represented by a series of nodes. Import the containing paragraph.
                    topNode = node.GetAncestor(NodeType.Paragraph);
                    break;
            }

            // If we have a node to import and it hasn't been added yet, schedule it.
            if (topNode != null && !imported.Contains(topNode))
            {
                nodesToImport.Add(topNode);
                imported.Add(topNode);
            }
        }

        // Import the collected nodes preserving their original order.
        foreach (Node srcNode in nodesToImport)
        {
            Node importedNode = importer.ImportNode(srcNode, true);
            dstBody.AppendChild(importedNode);
        }

        // Save the resulting document containing only tables, images, and fields.
        dstDoc.Save("output.docx");
    }
}
