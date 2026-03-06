using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk using the Document(string) constructor.
        Document doc = new Document("Input.docx");

        // Get a live collection of all Comment nodes in the document (including those in sub‑nodes).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Convert the NodeCollection to a strongly‑typed List<Comment>.
        List<Comment> comments = new List<Comment>();
        foreach (Comment comment in commentNodes)
        {
            comments.Add(comment);
        }

        // Example usage: iterate over the collected comments and output their details.
        foreach (Comment comment in comments)
        {
            // GetText() returns the comment's text including any child paragraphs.
            string text = comment.GetText().Trim();
            Console.WriteLine($"Author: {comment.Author}, Date: {comment.DateTime}, Text: {text}");
        }
    }
}
