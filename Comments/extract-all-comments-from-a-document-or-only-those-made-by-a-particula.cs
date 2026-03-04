using System;
using System.Collections.Generic;
using Aspose.Words;

class CommentExtractor
{
    // Extracts comments from a DOCX file.
    // If authorFilter is null or empty, all comments are returned.
    public static List<string> ExtractComments(string docPath, string authorFilter = null)
    {
        // Load the document from the file system (lifecycle rule).
        Document doc = new Document(docPath);

        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        List<string> comments = new List<string>();

        foreach (Comment comment in commentNodes)
        {
            // Apply author filter if provided.
            if (!string.IsNullOrEmpty(authorFilter) &&
                !string.Equals(comment.Author, authorFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            // Get the full text of the comment (including its child paragraphs).
            string text = comment.GetText().Trim();

            // Store comment information.
            comments.Add($"Author: {comment.Author}; Text: {text}");
        }

        return comments;
    }

    // Example usage.
    static void Main()
    {
        string filePath = @"C:\Docs\Sample.docx";

        // Extract all comments.
        List<string> allComments = ExtractComments(filePath);
        Console.WriteLine("All comments:");
        foreach (string c in allComments)
            Console.WriteLine(c);

        // Extract comments made by a specific author.
        List<string> johnComments = ExtractComments(filePath, "John Doe");
        Console.WriteLine("\nComments by John Doe:");
        foreach (string c in johnComments)
            Console.WriteLine(c);
    }
}
