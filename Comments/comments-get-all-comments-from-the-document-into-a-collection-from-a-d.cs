using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from file (lifecycle rule: load)
        Document doc = new Document("Input.docx");

        // Get all comment nodes in the document (including those in sub‑nodes)
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Convert the NodeCollection to a strongly‑typed List<Comment>
        List<Comment> comments = new List<Comment>();
        foreach (Comment comment in commentNodes)
        {
            comments.Add(comment);
        }

        // Example usage: print each comment's author and text
        foreach (Comment comment in comments)
        {
            Console.WriteLine($"Author: {comment.Author}, Text: {comment.GetText().Trim()}");
        }
    }
}
