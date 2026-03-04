using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file.
        string docPath = "input.docx";

        // Load the document using the provided Document constructor.
        Document doc = new Document(docPath);

        // Retrieve all comment nodes in the document (deep search).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Store the comments in a strongly‑typed collection.
        List<Comment> comments = new List<Comment>();

        // Iterate over the NodeCollection and add each Comment to the list.
        foreach (Comment comment in commentNodes)
        {
            comments.Add(comment);
        }

        // Example usage: print author and comment text.
        foreach (Comment comment in comments)
        {
            // GetText() returns the comment's content including any child paragraphs.
            string text = comment.GetText().Trim();
            Console.WriteLine($"Author: {comment.Author}, Text: {text}");
        }
    }
}
