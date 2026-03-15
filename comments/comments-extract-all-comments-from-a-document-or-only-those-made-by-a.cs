using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing; // Needed for NodeType enum

namespace AsposeWordsExamples
{
    /// <summary>
    /// Provides methods to extract comments from a DOCX document.
    /// </summary>
    public static class CommentExtractor
    {
        /// <summary>
        /// Loads the document from the specified file and returns the text of all comments.
        /// </summary>
        /// <param name="filePath">Full path to the DOCX file.</param>
        /// <returns>List of comment texts in the order they appear in the document.</returns>
        public static List<string> ExtractAllComments(string filePath)
        {
            // Load the document using the Document(string) constructor (lifecycle rule).
            Document doc = new Document(filePath);

            // Retrieve all Comment nodes in the document (including those inside headers/footers).
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            List<string> comments = new List<string>();
            foreach (Comment comment in commentNodes)
            {
                // The GetText method returns the full comment text (including any nested paragraphs/tables).
                comments.Add(comment.GetText().Trim());
            }

            return comments;
        }

        /// <summary>
        /// Loads the document from the specified file and returns the text of comments authored by the given author.
        /// </summary>
        /// <param name="filePath">Full path to the DOCX file.</param>
        /// <param name="author">Author name to filter comments (case‑sensitive).</param>
        /// <returns>List of comment texts authored by the specified person.</returns>
        public static List<string> ExtractCommentsByAuthor(string filePath, string author)
        {
            // Load the document using the Document(string) constructor (lifecycle rule).
            Document doc = new Document(filePath);

            // Retrieve all Comment nodes in the document.
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            List<string> comments = new List<string>();
            foreach (Comment comment in commentNodes)
            {
                // Filter by the Author property.
                if (string.Equals(comment.Author, author, StringComparison.Ordinal))
                {
                    comments.Add(comment.GetText().Trim());
                }
            }

            return comments;
        }

        // Example usage.
        public static void Main()
        {
            string docPath = @"C:\Docs\Sample.docx";

            // Extract every comment.
            List<string> allComments = ExtractAllComments(docPath);
            Console.WriteLine("All comments:");
            foreach (string txt in allComments)
                Console.WriteLine("- " + txt);

            // Extract only comments written by "John Doe".
            List<string> johnComments = ExtractCommentsByAuthor(docPath, "John Doe");
            Console.WriteLine("\nComments by John Doe:");
            foreach (string txt in johnComments)
                Console.WriteLine("- " + txt);
        }
    }
}
