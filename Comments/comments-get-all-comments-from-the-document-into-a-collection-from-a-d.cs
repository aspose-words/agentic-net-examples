using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve all comment nodes from the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Store the comments in a strongly‑typed collection.
        List<Comment> comments = new List<Comment>();
        foreach (Comment comment in commentNodes)
        {
            comments.Add(comment);
        }

        // Example usage: print each comment's author and text.
        foreach (Comment c in comments)
        {
            Console.WriteLine($"Author: {c.Author}, Text: {c.GetText()}");
        }
    }
}
