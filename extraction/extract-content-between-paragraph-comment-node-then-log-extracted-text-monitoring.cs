using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Get the body of the first section.
        Body body = doc.FirstSection.Body;

        // Find the first paragraph in the body.
        Paragraph paragraph = body.FirstParagraph;
        if (paragraph == null)
        {
            Console.WriteLine("No paragraph found in the document.");
            return;
        }

        // Locate the first comment node that appears after the paragraph.
        Node commentNode = null;
        for (Node node = paragraph.NextSibling; node != null; node = node.NextSibling)
        {
            if (node.NodeType == NodeType.Comment)
            {
                commentNode = node;
                break;
            }
        }

        if (commentNode == null)
        {
            Console.WriteLine("No comment node found after the paragraph.");
            return;
        }

        // Build a collection of all child nodes of the body.
        NodeCollection bodyChildren = body.GetChildNodes(NodeType.Any, true);

        // Determine the indices of the paragraph and the comment within the collection.
        int startIndex = bodyChildren.IndexOf(paragraph);
        int endIndex   = bodyChildren.IndexOf(commentNode);

        // Extract text from nodes that lie strictly between the paragraph and the comment.
        StringBuilder extractedText = new StringBuilder();
        for (int i = startIndex + 1; i < endIndex; i++)
        {
            Node current = bodyChildren[i];
            // GetText returns the visible text of the node (including runs, fields, etc.).
            extractedText.Append(current.GetText());
        }

        // Log the extracted text to the monitoring system (here we use Console as a placeholder).
        Console.WriteLine("Extracted text between paragraph and comment:");
        Console.WriteLine(extractedText.ToString());
    }
}
