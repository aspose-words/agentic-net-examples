using System;
using System.Collections.Generic;
using Aspose.Words;

class CommentProcessor
{
    // Extracts all comments from a DOCX file and returns them as a list of strings.
    public static List<string> ExtractComments(string inputPath)
    {
        // Load the document from the specified file.
        Document doc = new Document(inputPath);

        // Get all comment nodes in the document (including those in headers/footers).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        List<string> comments = new List<string>();
        foreach (Comment comment in commentNodes)
        {
            // The comment text is stored in the comment's child runs.
            // GetText() returns the concatenated text of the comment.
            comments.Add(comment.GetText().Trim());
        }

        return comments;
    }

    // Removes all comments from a DOCX file and saves the result to a new file.
    public static void RemoveComments(string inputPath, string outputPath)
    {
        // Load the document from the specified file.
        Document doc = new Document(inputPath);

        // Collect comment nodes first because removing them while iterating would modify the collection.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        List<Comment> commentsToRemove = new List<Comment>();
        foreach (Comment comment in commentNodes)
            commentsToRemove.Add(comment);

        // Remove each comment from its parent.
        foreach (Comment comment in commentsToRemove)
            comment.Remove();

        // Save the modified document.
        doc.Save(outputPath);
    }

    // Example usage.
    static void Main()
    {
        string sourceFile = @"C:\Docs\Sample.docx";

        // 1. Extract comments.
        List<string> extracted = ExtractComments(sourceFile);
        Console.WriteLine("Extracted Comments:");
        foreach (string txt in extracted)
            Console.WriteLine("- " + txt);

        // 2. Remove comments and save to a new file.
        string cleanedFile = @"C:\Docs\Sample_NoComments.docx";
        RemoveComments(sourceFile, cleanedFile);
        Console.WriteLine($"Comments removed and saved to: {cleanedFile}");
    }
}
