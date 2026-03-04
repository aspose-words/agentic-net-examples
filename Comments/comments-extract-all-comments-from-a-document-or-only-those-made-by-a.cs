using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class CommentExtractor
{
    /// <summary>
    /// Extracts comments from a DOCX file.
    /// If <paramref name="author"/> is null or empty, all comments are extracted;
    /// otherwise only comments authored by the specified person are extracted.
    /// The extracted comments are written to <paramref name="outputPath"/> as plain text.
    /// </summary>
    public static void ExtractComments(string inputPath, string outputPath, string author = null)
    {
        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Collect comment texts.
        List<string> extracted = new List<string>();

        // Retrieve all Comment nodes in the document (including those in headers/footers).
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

        foreach (Comment comment in comments)
        {
            // If an author filter is supplied, skip comments that do not match.
            if (!string.IsNullOrEmpty(author) && !string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                continue;

            // Build a readable representation of the comment.
            string text = $"Author: {comment.Author}{Environment.NewLine}" +
                          $"Date  : {comment.DateTime}{Environment.NewLine}" +
                          $"Text  : {comment.GetText().Trim()}{Environment.NewLine}" +
                          $"---";

            extracted.Add(text);
        }

        // Write the results to the output file.
        File.WriteAllLines(outputPath, extracted);
    }

    // Example usage.
    public static void Main()
    {
        string sourceDoc = @"C:\Docs\Sample.docx";
        string allCommentsOutput = @"C:\Docs\AllComments.txt";
        string authorCommentsOutput = @"C:\Docs\JohnDoeComments.txt";

        // Extract all comments.
        ExtractComments(sourceDoc, allCommentsOutput);

        // Extract only comments made by "John Doe".
        ExtractComments(sourceDoc, authorCommentsOutput, "John Doe");
    }
}
