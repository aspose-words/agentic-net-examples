using System;
using System.Collections.Generic;
using Aspose.Words;
using AsposeRange = Aspose.Words.Range;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file
        Document doc = new Document("InputDocument.docx");

        // Locate the start and end nodes for the desired range.
        // In this example we use the first and third paragraphs in the first section.
        Paragraph startNode = (Paragraph)doc.FirstSection.Body.GetChild(NodeType.Paragraph, 0, true);
        Paragraph endNode   = (Paragraph)doc.FirstSection.Body.GetChild(NodeType.Paragraph, 2, true);

        // Obtain the range that starts at the start node.
        AsposeRange subRange = startNode.Range;

        // Collect nodes from the start node up to (and including) the end node.
        List<Node> nodesInRange = new List<Node>();
        foreach (Node node in subRange)
        {
            nodesInRange.Add(node);
            if (node == endNode)
                break;
        }

        // Example: replace the text "OldText" with "NewText" in the selected range.
        foreach (Node node in nodesInRange)
        {
            if (node.NodeType == NodeType.Run)
            {
                Run run = (Run)node;
                run.Text = run.Text.Replace("OldText", "NewText");
            }
        }

        // Save the modified document
        doc.Save("OutputDocument.docx");
    }
}
