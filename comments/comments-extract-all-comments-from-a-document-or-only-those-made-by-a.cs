using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing; // Required for NodeType enum

namespace CommentExtractionDemo
{
    public static class CommentExtractor
    {
        // Extracts the text of all comments in the specified DOCX file.
        public static List<string> ExtractAllComments(string docxPath)
        {
            // Load the document from the file system.
            Document doc = new Document(docxPath);

            // Prepare a list to hold comment texts.
            List<string> comments = new List<string>();

            // Retrieve all comment nodes in the document (deep search).
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            // Iterate through each comment node and collect its author and text.
            foreach (Comment comment in commentNodes)
            {
                // Trim removes trailing paragraph breaks.
                string text = comment.GetText().Trim();
                comments.Add($"{comment.Author}: {text}");
            }

            return comments;
        }

        // Extracts the text of comments authored by the specified author.
        public static List<string> ExtractCommentsByAuthor(string docxPath, string authorName)
        {
            Document doc = new Document(docxPath);
            List<string> comments = new List<string>();
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in commentNodes)
            {
                if (string.Equals(comment.Author, authorName, StringComparison.OrdinalIgnoreCase))
                {
                    string text = comment.GetText().Trim();
                    comments.Add($"{comment.Author}: {text}");
                }
            }

            return comments;
        }

        // Example usage.
        public static void Main()
        {
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Get all comments.
            List<string> allComments = ExtractAllComments(inputPath);
            Console.WriteLine("All Comments:");
            foreach (string c in allComments)
                Console.WriteLine(c);

            // Get comments by a specific author.
            string targetAuthor = "John Doe";
            List<string> authorComments = ExtractCommentsByAuthor(inputPath, targetAuthor);
            Console.WriteLine($"\nComments by \"{targetAuthor}\":");
            foreach (string c in authorComments)
                Console.WriteLine(c);
        }
    }
}
