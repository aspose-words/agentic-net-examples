using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class CommentExtractor
{
    // Extracts comments from a DOCX file.
    // If author is null, returns all comments; otherwise returns only comments by the specified author.
    public static List<string> ExtractComments(string filePath, string author = null)
    {
        // Load the document from the specified path.
        Document doc = new Document(filePath);

        // Retrieve all comment nodes in the document (including nested comments).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        List<string> extracted = new List<string>();

        foreach (Comment comment in commentNodes)
        {
            // If an author filter is provided, skip comments that don't match.
            if (author != null && !string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                continue;

            // Get the plain text of the comment.
            string commentText = comment.ToString(SaveFormat.Text).Trim();

            // Build a readable representation.
            string entry = $"Author: {comment.Author}, Date: {comment.DateTime}, Text: {commentText}";
            extracted.Add(entry);
        }

        return extracted;
    }

    // Example usage.
    public static void Main()
    {
        string docPath = "Input.docx";

        // Extract all comments.
        List<string> allComments = ExtractComments(docPath);
        Console.WriteLine("All Comments:");
        foreach (string c in allComments)
            Console.WriteLine(c);

        // Extract comments made by a specific author.
        string targetAuthor = "John Doe";
        List<string> authorComments = ExtractComments(docPath, targetAuthor);
        Console.WriteLine($"\nComments by {targetAuthor}:");
        foreach (string c in authorComments)
            Console.WriteLine(c);
    }
}
