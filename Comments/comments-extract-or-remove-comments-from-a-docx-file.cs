using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Comparing;

class CommentProcessor
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // Extract all comments and write them to the console.
        // -----------------------------------------------------------------
        // Get all comment nodes in the document (including nested comments).
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

        Console.WriteLine("Extracted Comments:");
        foreach (Comment comment in comments)
        {
            // The comment text is stored in the comment's story (its paragraphs).
            string commentText = comment.ToString(SaveFormat.Text).Trim();
            Console.WriteLine($"- Author: {comment.Author}, Date: {comment.DateTime}");
            Console.WriteLine($"  Text: {commentText}");
        }

        // -----------------------------------------------------------------
        // Remove all comments from the document.
        // -----------------------------------------------------------------
        // Iterate backwards so that removal does not affect the collection indexing.
        for (int i = comments.Count - 1; i >= 0; i--)
        {
            comments[i].Remove();
        }

        // Save the document without comments.
        doc.Save("Output_NoComments.docx");
    }
}
