using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTX template file.
        Document doc = new Document("Template.dotx");

        // Example: define the start and end nodes.
        // Here we use the first and third paragraphs in the main body.
        // Adjust the indices as needed for your specific document.
        Paragraph startNode = doc.FirstSection.Body.Paragraphs[0];
        Paragraph endNode = doc.FirstSection.Body.Paragraphs[2];

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("Start or end node not found.");
            return;
        }

        // Extract the text that lies between the two nodes (exclusive).
        string betweenText = GetTextBetweenNodes(startNode, endNode);

        Console.WriteLine("Extracted text between the nodes:");
        Console.WriteLine(betweenText);
    }

    // Traverses the document in pre‑order and concatenates the text of all nodes
    // that appear after 'start' and before 'end'.
    static string GetTextBetweenNodes(Node start, Node end)
    {
        // Move to the node that follows the start node.
        Node current = start.NextPreOrder(start.Document);
        var parts = new List<string>();

        // Continue until we reach the end node or the document ends.
        while (current != null && current != end)
        {
            // Append the textual representation of the current node.
            parts.Add(current.GetText());

            // Advance to the next node in pre‑order traversal.
            current = current.NextPreOrder(current.Document);
        }

        // Combine all collected parts into a single string.
        return string.Concat(parts);
    }
}
