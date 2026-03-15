using System;
using System.Collections.Generic;
using Aspose.Words;

namespace AsposeWordsCommentExtraction
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string sourceFile = @"C:\Docs\SampleDocument.docx";

            // Extract all comments.
            List<string> allComments = ExtractComments(sourceFile);
            Console.WriteLine("All comments:");
            foreach (string comment in allComments)
                Console.WriteLine("- " + comment);

            // Extract only comments authored by "John Doe".
            const string author = "John Doe";
            List<string> johnComments = ExtractCommentsByAuthor(sourceFile, author);
            Console.WriteLine($"\nComments by \"{author}\":");
            foreach (string comment in johnComments)
                Console.WriteLine("- " + comment);
        }

        /// <summary>
        /// Loads a document and returns the text of every comment it contains.
        /// </summary>
        /// <param name="filePath">Full path to the DOCX file.</param>
        /// <returns>List of comment texts.</returns>
        static List<string> ExtractComments(string filePath)
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Collect comment nodes.
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
            List<string> comments = new List<string>(commentNodes.Count);

            foreach (Comment comment in commentNodes)
            {
                // The Range property contains the full comment content.
                // Trim to remove leading/trailing whitespace.
                comments.Add(comment.Range.Text.Trim());
            }

            return comments;
        }

        /// <summary>
        /// Loads a document and returns the text of comments authored by the specified person.
        /// </summary>
        /// <param name="filePath">Full path to the DOCX file.</param>
        /// <param name="author">Author name to filter by (case‑sensitive).</param>
        /// <returns>List of comment texts authored by the given person.</returns>
        static List<string> ExtractCommentsByAuthor(string filePath, string author)
        {
            Document doc = new Document(filePath);
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
            List<string> comments = new List<string>(commentNodes.Count);

            foreach (Comment comment in commentNodes)
            {
                // Compare the Author property; ignore comments with a different author.
                if (string.Equals(comment.Author, author, StringComparison.Ordinal))
                {
                    comments.Add(comment.Range.Text.Trim());
                }
            }

            return comments;
        }
    }
}
