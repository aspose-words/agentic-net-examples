using System;
using System.Collections.Generic;
using Aspose.Words;

class RemoveCommentsExample
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Path to the output DOCX file.
        string outputPath = @"C:\Docs\output.docx";

        // Optional: specify an author name to remove only that author's comments.
        // Set to null or empty string to remove all comments.
        string authorToRemove = "John Doe";

        // Load the document.
        Document doc = new Document(inputPath);

        // Collect comments that need to be removed.
        List<Comment> commentsToRemove = new List<Comment>();

        // Get all comment nodes in the document (including nested comments).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment comment in commentNodes)
        {
            // If an author is specified, remove only comments from that author.
            // Otherwise, remove every comment.
            if (string.IsNullOrEmpty(authorToRemove) || comment.Author == authorToRemove)
                commentsToRemove.Add(comment);
        }

        // Remove the collected comments from the document.
        foreach (Comment comment in commentsToRemove)
            comment.Remove();

        // Save the modified document.
        doc.Save(outputPath);
    }
}
