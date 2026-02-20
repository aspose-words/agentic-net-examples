using System;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // ------------------------------------------------------------
        // Example 1: Extract text from a plain (non‑ranged) content control
        // ------------------------------------------------------------
        // Find the content control (structured document tag) by its title.
        StructuredDocumentTag plainTag = doc.Range.StructuredDocumentTags.GetByTitle("PlainTag") as StructuredDocumentTag;

        if (plainTag != null)
        {
            // The GetText() method returns the text that the content control contains.
            string plainContent = plainTag.GetText();
            Console.WriteLine("Plain content control text:");
            Console.WriteLine(plainContent);
        }

        // ------------------------------------------------------------
        // Example 2: Extract text from a ranged content control
        // ------------------------------------------------------------
        // Ranged tags consist of a start node and an end node. We locate the start node
        // by its title (or any other criteria) and then find the matching end node by Id.
        var rangeStarts = doc.GetChildNodes(NodeType.StructuredDocumentTagRangeStart, true)
                             .Cast<StructuredDocumentTagRangeStart>()
                             .Where(s => s.Title == "RangedTag");

        foreach (var startNode in rangeStarts)
        {
            // Find the corresponding end node with the same Id.
            var endNode = doc.GetChildNodes(NodeType.StructuredDocumentTagRangeEnd, true)
                             .Cast<StructuredDocumentTagRangeEnd>()
                             .FirstOrDefault(e => e.Id == startNode.Id);

            if (endNode != null)
            {
                string rangedContent = GetTextBetween(startNode, endNode);
                Console.WriteLine("Ranged content control text:");
                Console.WriteLine(rangedContent);
            }
        }
    }

    // Helper method that walks the document tree from the node after 'start'
    // up to (but not including) the 'end' node, concatenating the text of each node.
    static string GetTextBetween(Node start, Node end)
    {
        StringBuilder sb = new StringBuilder();

        // Move to the first node after the start node in a pre‑order traversal.
        Node cur = start.NextPreOrder(start);

        // Continue until we reach the end node or the document ends.
        while (cur != null && cur != end)
        {
            sb.Append(cur.GetText());
            cur = cur.NextPreOrder(cur);
        }

        return sb.ToString();
    }
}
