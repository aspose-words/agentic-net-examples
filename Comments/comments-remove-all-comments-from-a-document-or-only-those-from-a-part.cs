using System;
using System.Collections.Generic;
using Aspose.Words;

namespace CommentRemovalExample
{
    public class CommentRemover
    {
        /// <summary>
        /// Removes all comments from a DOCX file, or only those authored by a specific user.
        /// </summary>
        /// <param name="inputFile">Path to the source DOCX document.</param>
        /// <param name="outputFile">Path where the cleaned document will be saved.</param>
        /// <param name="author">
        /// Optional author name. If null or empty, all comments are removed.
        /// If provided, only comments whose Author property matches this value (case‑insensitive) are removed.
        /// </param>
        public static void RemoveComments(string inputFile, string outputFile, string? author = null)
        {
            // Load the document using the provided constructor rule.
            Document doc = new Document(inputFile);

            // Retrieve all comment nodes in the document (deep search).
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Collect nodes to delete to avoid modifying the collection while iterating.
            List<Comment> toDelete = new List<Comment>();

            foreach (Comment comment in comments)
            {
                // If an author filter is supplied, compare case‑insensitively.
                if (!string.IsNullOrEmpty(author))
                {
                    if (string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                    {
                        toDelete.Add(comment);
                    }
                }
                else
                {
                    // No author filter – mark every comment for removal.
                    toDelete.Add(comment);
                }
            }

            // Remove the selected comments from the document.
            foreach (Comment comment in toDelete)
            {
                comment.Remove();
            }

            // Save the modified document using the provided Save rule.
            doc.Save(outputFile);
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point for the example.
        /// Usage: dotnet run <inputFile> <outputFile> [author]
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <inputFile> <outputFile> [author]");
                return;
            }

            string inputFile = args[0];
            string outputFile = args[1];
            string? author = args.Length >= 3 ? args[2] : null;

            try
            {
                CommentRemover.RemoveComments(inputFile, outputFile, author);
                Console.WriteLine($"Comments {(author == null ? "all" : $"by '{author}'")} removed successfully. Output saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
