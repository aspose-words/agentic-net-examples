using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

public class CommentExtractor
{
    /// <summary>
    /// Extracts comment texts from a DOCX file.
    /// If <paramref name="author"/> is null or empty, returns all comments.
    /// Otherwise returns only comments authored by the specified person.
    /// </summary>
    /// <param name="filePath">Path to the DOCX file.</param>
    /// <param name="author">Optional author name to filter comments.</param>
    /// <returns>List of comment texts.</returns>
    public static List<string> ExtractComments(string filePath, string author = null)
    {
        // Load the document using the provided constructor (lifecycle rule).
        Document doc = new Document(filePath);

        // Collect all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        List<string> result = new List<string>();

        foreach (Comment comment in commentNodes)
        {
            // If an author filter is supplied, skip comments that don't match.
            if (!string.IsNullOrEmpty(author) && !string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                continue;

            // The comment text is stored in the comment's range.
            // Trim to remove leading/trailing whitespace and line breaks.
            string text = comment.GetText().Trim();

            // Optionally include author information in the output.
            result.Add($"Author: {comment.Author}\nText: {text}");
        }

        return result;
    }

    // Example usage.
    public static void Main()
    {
        string docPath = @"C:\Docs\Sample.docx";

        // Extract all comments.
        List<string> allComments = ExtractComments(docPath);
        Console.WriteLine("All Comments:");
        foreach (string c in allComments)
            Console.WriteLine(c + "\n---");

        // Extract only comments made by "John Doe".
        List<string> johnComments = ExtractComments(docPath, "John Doe");
        Console.WriteLine("Comments by John Doe:");
        foreach (string c in johnComments)
            Console.WriteLine(c + "\n---");
    }
}
