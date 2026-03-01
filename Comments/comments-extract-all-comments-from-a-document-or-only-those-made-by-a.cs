using System;
using System.Collections.Generic;
using Aspose.Words;

class CommentExtractor
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string inputPath = "InputDocument.docx";

        // Load the source document (lifecycle rule: load).
        Document sourceDoc = new Document(inputPath);

        // Retrieve all comment nodes in the document (including those in headers/footers).
        NodeCollection allComments = sourceDoc.GetChildNodes(NodeType.Comment, true);

        // Example 1: Extract all comments.
        List<Comment> extractedAll = new List<Comment>();
        foreach (Comment comment in allComments)
        {
            extractedAll.Add(comment);
        }

        // Example 2: Extract comments made by a specific author.
        const string targetAuthor = "John Doe"; // Set to null or empty to skip author filtering.
        List<Comment> extractedByAuthor = new List<Comment>();
        if (!string.IsNullOrEmpty(targetAuthor))
        {
            foreach (Comment comment in allComments)
            {
                if (string.Equals(comment.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
                {
                    extractedByAuthor.Add(comment);
                }
            }
        }

        // Create a new document to store the extracted comments (lifecycle rule: create).
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);

        // Helper method to write a collection of comments into the result document.
        void WriteComments(IEnumerable<Comment> comments, string heading)
        {
            builder.Writeln(heading);
            builder.Font.Bold = true;
            builder.Writeln($"Total: {System.Linq.Enumerable.Count(comments)}");
            builder.Font.Bold = false;
            builder.Writeln();

            foreach (Comment c in comments)
            {
                // Write author and comment text.
                builder.Writeln($"Author: {c.Author}");
                // The comment text is stored in its child paragraphs; convert to plain text.
                string commentText = c.ToString(SaveFormat.Text).Trim();
                builder.Writeln($"Comment: {commentText}");
                builder.Writeln(); // Blank line between comments.
            }

            builder.Writeln(new string('-', 40));
            builder.Writeln();
        }

        // Write all comments.
        WriteComments(extractedAll, "All Comments");

        // Write filtered comments if a target author was specified.
        if (!string.IsNullOrEmpty(targetAuthor))
        {
            WriteComments(extractedByAuthor, $"Comments by \"{targetAuthor}\"");
        }

        // Save the result document containing the extracted comments (lifecycle rule: save).
        const string outputPath = "ExtractedComments.docx";
        resultDoc.Save(outputPath);
    }
}
