using System;
using System.Collections.Generic;
using Aspose.Words;

namespace CommentExtractionDemo
{
    public static class CommentExtractor
    {
        /// <summary>
        /// Extracts comment texts from a DOCX file.
        /// If <paramref name="author"/> is null or empty, returns all comments;
        /// otherwise returns only comments authored by the specified person.
        /// </summary>
        /// <param name="filePath">Path to the DOCX file.</param>
        /// <param name="author">Optional author name to filter comments.</param>
        /// <returns>List of comment texts.</returns>
        public static List<string> ExtractComments(string filePath, string? author = null)
        {
            // Load the document (lifecycle rule: use Document constructor)
            Document doc = new Document(filePath);

            // Prepare the result list
            List<string> comments = new List<string>();

            // Retrieve all comment nodes in the document (including those inside footnotes, headers, etc.)
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in commentNodes)
            {
                // If an author filter is supplied, skip comments that don't match
                if (!string.IsNullOrEmpty(author) &&
                    !string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Get the full text of the comment (including any paragraphs/tables it may contain)
                // Trim trailing line breaks that Aspose adds to comment text
                string commentText = comment.GetText()?.TrimEnd('\r', '\n') ?? string.Empty;

                comments.Add(commentText);
            }

            return comments;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: CommentExtractionDemo <docxPath> [author]");
                return;
            }

            string path = args[0];
            string? author = args.Length > 1 ? args[1] : null;

            List<string> extracted = CommentExtractor.ExtractComments(path, author);
            Console.WriteLine($"Found {extracted.Count} comment(s):");
            foreach (var c in extracted)
                Console.WriteLine("- " + c);
        }
    }
}
