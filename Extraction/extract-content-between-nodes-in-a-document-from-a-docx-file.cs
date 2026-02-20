using System;
using System.Text;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("Input.docx");

        // Identify the start and end nodes.
        // For example, get the 3rd paragraph as the start node and the 6th paragraph as the end node.
        // Adjust the logic as needed (e.g., use bookmarks, headings, etc.).
        Node startNode = doc.GetChild(NodeType.Paragraph, 2, true); // zero‑based index
        Node endNode   = doc.GetChild(NodeType.Paragraph, 5, true);

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("Start or end node not found.");
            return;
        }

        // Collect text from the start node up to and including the end node.
        StringBuilder extractedText = new StringBuilder();

        for (Node cur = startNode; cur != null; cur = cur.NextSibling)
        {
            extractedText.Append(cur.GetText());

            // Stop after processing the end node.
            if (cur == endNode)
                break;
        }

        // Output the extracted content.
        Console.WriteLine("Extracted text between nodes:");
        Console.WriteLine(extractedText.ToString());

        // Optionally, save the extracted text to a file.
        System.IO.File.WriteAllText("ExtractedContent.txt", extractedText.ToString());
    }
}
