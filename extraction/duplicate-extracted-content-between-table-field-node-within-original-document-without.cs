using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

class DuplicateContentBetweenTableAndField
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Locate the first table in the document.
        Table table = doc.GetChildNodes(NodeType.Table, true)
                         .Cast<Table>()
                         .FirstOrDefault();

        // Locate the first field (any field) in the document.
        FieldStart fieldStart = doc.GetChildNodes(NodeType.FieldStart, true)
                                   .Cast<FieldStart>()
                                   .FirstOrDefault();

        if (table == null || fieldStart == null)
        {
            Console.WriteLine("Required nodes not found.");
            return;
        }

        // Ensure the field appears after the table in the document order.
        // If not, swap the references so that we always work from earlier to later node.
        Node startNode = table;
        Node endNode = fieldStart;
        if (IsNodeAfter(startNode, endNode))
        {
            // field is before table – swap to keep startNode before endNode.
            startNode = fieldStart;
            endNode = table;
        }

        // Collect all nodes that are strictly between startNode and endNode.
        List<Node> nodesBetween = new List<Node>();
        Node cur = startNode.NextPreOrder(doc);
        while (cur != null && cur != endNode)
        {
            nodesBetween.Add(cur);
            cur = cur.NextPreOrder(doc);
        }

        // Prepare a NodeImporter that will keep source formatting.
        NodeImporter importer = new NodeImporter(doc, doc, ImportFormatMode.KeepSourceFormatting);

        // Determine the insertion point – after the field's end node.
        Field field = fieldStart.GetField();
        Node insertionPoint = field.End; // FieldEnd node

        // Duplicate each collected node and insert after the insertion point.
        foreach (Node srcNode in nodesBetween)
        {
            // Import the node (deep clone) preserving formatting.
            Node importedNode = importer.ImportNode(srcNode, true);

            // Insert the imported node after the current insertion point.
            // After each insertion, move the insertion point forward.
            insertionPoint.ParentNode.InsertAfter(importedNode, insertionPoint);
            insertionPoint = importedNode;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Helper method to determine if nodeA appears after nodeB in the document order.
    private static bool IsNodeAfter(Node nodeA, Node nodeB)
    {
        // Walk forward from nodeB; if we encounter nodeA, then nodeA is after nodeB.
        Node cur = nodeB;
        while (cur != null)
        {
            if (cur == nodeA)
                return true;
            cur = cur.NextPreOrder(nodeA.Document);
        }
        return false;
    }
}
