using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

class RangeBetweenNodesExample
{
    static void Main()
    {
        // Load the input document (replace with your actual file path).
        Document doc = new Document(@"Input.docx");

        // Locate the start and end paragraphs (first and third paragraphs in this example).
        Paragraph startParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        Paragraph endParagraph   = (Paragraph)doc.GetChild(NodeType.Paragraph, 2, true);

        // Collect all nodes that lie between the start and end nodes, inclusive.
        List<Node> nodesInRange = new List<Node>();
        bool inside = false;
        foreach (Node node in doc)
        {
            if (node == startParagraph)
                inside = true;

            if (inside)
                nodesInRange.Add(node);

            if (node == endParagraph)
                break;
        }

        // Build a temporary document that contains only the collected nodes.
        Document subDoc = new Document();
        subDoc.EnsureMinimum(); // Guarantees a section and a body.

        NodeImporter importer = new NodeImporter(doc, subDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Node node in nodesInRange)
        {
            Node imported = importer.ImportNode(node, true);
            subDoc.FirstSection.Body.AppendChild(imported);
        }

        // The Range of the temporary document now represents the range that starts at
        // startParagraph and ends at endParagraph.
        Aspose.Words.Range desiredRange = subDoc.Range;

        // Example usage: output the text of the range.
        Console.WriteLine("Extracted range text:");
        Console.WriteLine(desiredRange.Text.Trim());

        // If you need to delete the range from the original document:
        foreach (Node node in nodesInRange)
        {
            node.Remove();
        }

        // Save the modified original document.
        doc.Save(@"Modified.docx");
    }
}
