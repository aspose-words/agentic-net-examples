using System;
using System.IO;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Find the start and end nodes.
        // For this example we look for paragraphs that contain the markers "[START]" and "[END]".
        Paragraph startNode = null;
        Paragraph endNode = null;

        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            string text = para.GetText();

            if (startNode == null && text.Contains("[START]"))
                startNode = para;

            if (endNode == null && text.Contains("[END]"))
                endNode = para;

            if (startNode != null && endNode != null)
                break;
        }

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("Start or end marker not found.");
            return;
        }

        // Collect text of all nodes that are between the start and end markers (exclusive).
        Node current = startNode.NextSibling;
        var extractedText = new System.Text.StringBuilder();

        while (current != null && current != endNode)
        {
            // Append the text of the current node.
            extractedText.Append(current.GetText());

            // Move to the next sibling; if null, walk up the tree to find the next node in document order.
            if (current.NextSibling != null)
                current = current.NextSibling;
            else
                current = current.ParentNode?.NextSibling;
        }

        // Output the extracted content.
        string result = extractedText.ToString().Trim();
        Console.WriteLine("Extracted content between markers:");
        Console.WriteLine(result);

        // Optionally, save the extracted text to a plain‑text file.
        File.WriteAllText("ExtractedContent.txt", result);
    }
}
