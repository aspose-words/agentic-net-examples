using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the RTF document using the provided load rule (RtfLoadOptions).
        var loadOptions = new RtfLoadOptions();
        Document doc = new Document("input.rtf", loadOptions);

        // ---------------------------------------------------------------------
        // 1. Extract the inner text of all ranged content controls (SDT range start).
        // ---------------------------------------------------------------------
        foreach (StructuredDocumentTagRangeStart sdtStart in doc.GetChildNodes(NodeType.StructuredDocumentTagRangeStart, true))
        {
            // The Range property of the start node represents the content inside the control.
            string innerText = sdtStart.Range.Text;

            Console.WriteLine($"Content Control Title: {sdtStart.Title}");
            Console.WriteLine("Extracted Text:");
            Console.WriteLine(innerText);
            Console.WriteLine(new string('-', 40));
        }

        // ---------------------------------------------------------------------
        // 2. Extract text that lies between two arbitrary nodes (e.g., two paragraphs).
        // ---------------------------------------------------------------------
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        Paragraph secondParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);

        string betweenText = GetTextBetween(firstParagraph, secondParagraph);

        Console.WriteLine("Text between first and second paragraph:");
        Console.WriteLine(betweenText);
    }

    /// <summary>
    /// Returns the concatenated text of all nodes that are positioned after <paramref name="startNode"/>
    /// and before <paramref name="endNode"/> (exclusive).
    /// </summary>
    private static string GetTextBetween(Node startNode, Node endNode)
    {
        if (startNode == null) throw new ArgumentNullException(nameof(startNode));
        if (endNode == null) throw new ArgumentNullException(nameof(endNode));
        if (startNode.Document != endNode.Document) throw new ArgumentException("Nodes must belong to the same document.");

        var sb = new StringBuilder();
        // Walk the sibling chain from the node that follows startNode up to (but not including) endNode.
        for (Node cur = startNode.NextSibling; cur != null && cur != endNode; cur = cur.NextSibling)
        {
            sb.Append(cur.GetText());
        }
        return sb.ToString();
    }
}
