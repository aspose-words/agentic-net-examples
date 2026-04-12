using System;
using System.Collections.Generic;
using Aspose.Words;

public class ExtractBetweenParagraphAndComment
{
    public static void Main()
    {
        // Create a new document and a builder for adding content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Insert a comment after the third paragraph.
        Comment comment = new Comment(doc)
        {
            Author = "Monitor",
            Initial = "M",
            DateTime = DateTime.Now
        };
        // Append the comment to the current paragraph (the third one).
        builder.CurrentParagraph.AppendChild(comment);

        // Add another paragraph after the comment.
        builder.Writeln("Paragraph after comment.");

        // Locate the start paragraph (first paragraph) and the comment node.
        Paragraph startParagraph = doc.FirstSection.Body.Paragraphs[0];
        Node commentNode = doc.GetChildNodes(NodeType.Comment, true)[0];

        // Extract text from nodes that appear between the start paragraph and the comment.
        List<string> extractedParts = new List<string>();
        Node current = startParagraph.NextSibling;
        while (current != null && current != commentNode)
        {
            if (current.NodeType == NodeType.Paragraph)
            {
                // Get the full text of the paragraph, including its ending character.
                extractedParts.Add(((Paragraph)current).GetText());
            }
            else if (current.NodeType == NodeType.Run)
            {
                extractedParts.Add(((Run)current).GetText());
            }
            // Move to the next sibling node.
            current = current.NextSibling;
        }

        // Combine the extracted parts into a single string.
        string extractedText = string.Join(string.Empty, extractedParts).Trim();

        // Log the extracted text to the monitoring system (simulated by console output).
        Console.WriteLine("Extracted text between paragraph and comment:");
        Console.WriteLine(extractedText);
    }
}
