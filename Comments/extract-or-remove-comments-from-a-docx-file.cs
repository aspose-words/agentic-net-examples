using System;
using Aspose.Words;

class CommentProcessor
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Path where the document without comments will be saved.
        string destinationPath = @"C:\Docs\DocumentWithoutComments.docx";

        // Load the document.
        Document doc = new Document(sourcePath);

        // Extract all comments.
        ExtractComments(doc);

        // Remove all comments from the document.
        RemoveAllComments(doc);

        // Save the modified document.
        doc.Save(destinationPath);
    }

    // Prints the text of each comment to the console.
    private static void ExtractComments(Document doc)
    {
        // Retrieve all comment nodes in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        Console.WriteLine("Extracted Comments:");
        foreach (Comment comment in commentNodes)
        {
            // The comment text is stored in the comment's child runs.
            string commentText = comment.GetText().Trim();
            Console.WriteLine($"- Author: {comment.Author}, Date: {comment.DateTime}");
            Console.WriteLine($"  Text: {commentText}");
        }
    }

    // Removes every comment node from the document.
    private static void RemoveAllComments(Document doc)
    {
        // Retrieve all comment nodes.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Remove comments starting from the last node to avoid collection modification issues.
        for (int i = commentNodes.Count - 1; i >= 0; i--)
        {
            commentNodes[i].Remove();
        }
    }
}
