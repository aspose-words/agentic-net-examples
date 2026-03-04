using System;
using System.Collections.Generic;
using Aspose.Words;

public class CommentExtractor
{
    /// <summary>
    /// Extracts comments from a DOCX file.
    /// If <paramref name="author"/> is null or empty, all comments are returned.
    /// Otherwise only comments authored by the specified person are returned.
    /// </summary>
    /// <param name="filePath">Path to the DOCX file.</param>
    /// <param name="author">Optional author name to filter comments.</param>
    /// <returns>List of comment texts (including author information).</returns>
    public static List<string> ExtractComments(string filePath, string author = null)
    {
        // Load the document using the provided constructor rule.
        Document doc = new Document(filePath);

        // Retrieve all comment nodes in the document (deep search).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        List<string> result = new List<string>();

        foreach (Comment comment in commentNodes)
        {
            // If an author filter is supplied, skip comments that don't match.
            if (!string.IsNullOrEmpty(author) &&
                !string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                continue;

            // Combine author and comment text for clarity.
            string formatted = $"Author: {comment.Author ?? "(no author)"}\nComment: {comment.GetText().Trim()}";
            result.Add(formatted);
        }

        return result;
    }

    // Example usage.
    public static void Main()
    {
        // Use a verbatim string for the file path to avoid escaping backslashes.
        string docPath = @"C:\Docs\Sample.docx";

        // Extract all comments.
        List<string> allComments = ExtractComments(docPath);
        Console.WriteLine("All comments:");
        foreach (string c in allComments)
            Console.WriteLine(c + "\n---");

        // Extract only comments made by "John Doe".
        List<string> johnComments = ExtractComments(docPath, "John Doe");
        Console.WriteLine("Comments by John Doe:");
        foreach (string c in johnComments)
            Console.WriteLine(c + "\n---");
    }
}
