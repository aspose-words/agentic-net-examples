using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main(string[] args)
    {
        // Path to the DOCX file.
        string filePath = "input.docx";

        // Author whose comments we want to extract.
        string author = "John Doe";

        // Load the document.
        Document doc = new Document(filePath);

        // Retrieve all comments made by the specified author.
        List<string> authorComments = GetCommentsByAuthor(doc, author);

        // Output the extracted comment texts.
        foreach (string commentText in authorComments)
        {
            Console.WriteLine(commentText);
        }
    }

    // Returns a list of comment texts authored by the given author.
    static List<string> GetCommentsByAuthor(Document doc, string author)
    {
        var comments = new List<string>();

        // Get all comment nodes in the document (including nested comments).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        foreach (Comment comment in commentNodes)
        {
            if (comment.Author == author)
            {
                // GetText() returns the full text of the comment, including its replies.
                comments.Add(comment.GetText());
            }
        }

        return comments;
    }
}
